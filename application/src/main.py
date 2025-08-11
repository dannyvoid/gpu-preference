#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
import sys
import ctypes
from dataclasses import dataclass
from enum import IntEnum
from pathlib import Path
from typing import Iterable, List, Optional, Dict

import psutil
import winreg

from PySide6.QtCore import Qt, QSize, QSortFilterProxyModel, QSettings
from PySide6.QtGui import (
    QAction,
    QActionGroup,
    QColor,
    QGuiApplication,
    QIcon,
    QPalette,
    QStandardItem,
    QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHeaderView,
    QInputDialog,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QStyledItemDelegate,
    QTableView,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QHBoxLayout,
    QSizePolicy,
    QStyle,
    QStyleOptionViewItem,
)

REG_PATH = r"Software\Microsoft\DirectX\UserGpuPreferences"


def normalize_path(p: str | Path) -> str:
    p = os.path.normpath(str(p))
    if len(p) >= 2 and p[1] == ":":
        p = p[0].upper() + p[1:]
    return p


def is_exe(path: str) -> bool:
    return os.path.isabs(path) and path.lower().endswith(".exe")


def contrasting_text_color(bg: QColor) -> QColor:
    luma = (0.299 * bg.red() + 0.587 * bg.green() + 0.114 * bg.blue()) / 255
    return QColor(0, 0, 0) if luma > 0.5 else QColor(255, 255, 255)


class Pref(IntEnum):
    POWER = 1
    PERF = 2

    @staticmethod
    def from_reg_value(val: str) -> "Pref":
        return Pref.PERF if "GpuPreference=2" in val else Pref.POWER


@dataclass
class Entry:
    exe: str
    pref: Pref


class RegistryManager:
    KEY_FLAGS = winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY

    @staticmethod
    def read_all() -> List[Entry]:
        out: List[Entry] = []
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
                i = 0
                while True:
                    name, val, _ = winreg.EnumValue(key, i)
                    out.append(
                        Entry(exe=normalize_path(name), pref=Pref.from_reg_value(val))
                    )
                    i += 1
        except OSError:
            pass
        return out

    @staticmethod
    def set_pref(exe_path: str, pref: Pref) -> None:
        exe_norm = normalize_path(exe_path)
        with winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER, REG_PATH, 0, RegistryManager.KEY_FLAGS
        ) as key:
            winreg.SetValueEx(
                key, exe_norm, 0, winreg.REG_SZ, f"GpuPreference={int(pref)};"
            )

    @staticmethod
    def delete(exe_path: str) -> None:
        exe_norm = normalize_path(exe_path)
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, REG_PATH, 0, RegistryManager.KEY_FLAGS
            ) as key:
                winreg.DeleteValue(key, exe_norm)
        except FileNotFoundError:
            pass

    @staticmethod
    def backup(to_file: str) -> None:
        data = {e.exe: int(e.pref) for e in RegistryManager.read_all()}
        with open(to_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    @staticmethod
    def restore(from_file: str) -> None:
        with open(from_file, "r", encoding="utf-8") as f:
            data: Dict[str, int] = json.load(f)
        for e in RegistryManager.read_all():
            RegistryManager.delete(e.exe)
        for exe, code in data.items():
            RegistryManager.set_pref(exe, Pref(code))


class PathDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return QLineEdit(parent)

    def setEditorData(self, editor: QLineEdit, index):
        editor.setText(index.model().data(index, Qt.EditRole))

    def setModelData(self, editor: QLineEdit, model: QStandardItemModel, index):
        new_path = normalize_path(editor.text())
        row = index.row()

        old_path = model.item(row, MainWindow.Columns.EXEC).text()
        if new_path != old_path:
            pref_text = model.item(row, MainWindow.Columns.PREF).text()
            high_perf = pref_text == model.parent().label_perf
            RegistryManager.delete(old_path)
            RegistryManager.set_pref(new_path, Pref.PERF if high_perf else Pref.POWER)
            model.setData(index, new_path, Qt.EditRole)

        model.parent()._update_row_existence(row)


class PrefDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        cb = QComboBox(parent)
        cb.setEditable(False)
        cb.setInsertPolicy(QComboBox.NoInsert)
        mw: MainWindow = index.model().parent()
        cb.addItems([mw.label_power, mw.label_perf])
        return cb

    def setEditorData(self, editor: QComboBox, index):
        mw: MainWindow = index.model().parent()
        editor.setCurrentIndex(
            0 if index.model().data(index, Qt.EditRole) == mw.label_power else 1
        )

    def setModelData(self, editor: QComboBox, model: QStandardItemModel, index):
        mw: MainWindow = model.parent()
        text = editor.currentText()
        exe = model.item(index.row(), MainWindow.Columns.EXEC).text()
        RegistryManager.set_pref(
            exe, Pref.PERF if editor.currentIndex() == 1 else Pref.POWER
        )
        model.setData(index, text, Qt.EditRole)
        mw.statusBar().showMessage(f"Updated GPU preference for {Path(exe).name}", 2000)


def set_role(btn: QToolButton, role: str):
    """Set a style role property that QSS can target, e.g., [role="primary"]"""
    btn.setProperty("role", role)
    btn.style().unpolish(btn)
    btn.style().polish(btn)
    btn.update()


def mark_checked(btn: QToolButton, checked: bool):
    """Mark button as checked for QSS selector [checked="true"]"""
    btn.setProperty("checked", "true" if checked else "false")
    btn.style().unpolish(btn)
    btn.style().polish(btn)
    btn.update()


class SegmentedControl(QWidget):
    def __init__(self, labels: List[str], parent=None):
        super().__init__(parent)
        self.setProperty("segmented", "true")
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.buttons: List[QToolButton] = []
        for txt in labels:
            b = QToolButton(self)
            b.setText(txt)
            b.setCheckable(True)
            b.setAutoExclusive(True)
            outer.addWidget(b)
            self.buttons.append(b)

    def on_change(self, idx: int, func):
        if 0 <= idx < len(self.buttons):
            self.buttons[idx].clicked.connect(func)

    def set_checked(self, idx: Optional[int]):
        """If idx is None, clear all; else check idx."""
        if idx is None:
            for b in self.buttons:
                b.setAutoExclusive(False)
                b.setChecked(False)
                mark_checked(b, False)
            for b in self.buttons:
                b.setAutoExclusive(True)
            return
        for i, b in enumerate(self.buttons):
            b.setChecked(i == idx)
            mark_checked(b, i == idx)

    def set_texts(self, labels: List[str]):
        for i, (b, t) in enumerate(zip(self.buttons, labels)):
            b.setText(t)


class RunningProcessDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add from Running Processes")
        self.resize(640, 420)

        layout = QVBoxLayout(self)
        self.search = QLineEdit(self, placeholderText="Search executables…")
        layout.addWidget(self.search)

        self.src_model = QStandardItemModel(0, 3, self)
        self.src_model.setHeaderData(0, Qt.Horizontal, "")
        self.src_model.setHeaderData(1, Qt.Horizontal, "Executable")
        self.src_model.setHeaderData(2, Qt.Horizontal, "PID")

        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setSourceModel(self.src_model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setFilterKeyColumn(1)

        self.table = QTableView(self)
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().hide()

        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)

        layout.addWidget(self.table)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._populate()
        self.search.textChanged.connect(self.proxy.setFilterFixedString)

    def _populate(self):
        seen: set[str] = set()
        for proc in psutil.process_iter(["exe", "pid"]):
            path = proc.info["exe"]
            if not path:
                continue
            path = normalize_path(path)
            if not is_exe(path) or path in seen:
                continue
            seen.add(path)

            chk = QStandardItem()
            chk.setCheckable(True)
            chk.setCheckState(Qt.Unchecked)

            exe_item = QStandardItem(Path(path).name)
            exe_item.setToolTip(path)

            pid_item = QStandardItem(str(proc.info["pid"]))
            self.src_model.appendRow([chk, exe_item, pid_item])

    def selected_paths(self) -> List[str]:
        out = []
        for row in range(self.src_model.rowCount()):
            if self.src_model.item(row, 0).checkState() == Qt.Checked:
                out.append(self.src_model.item(row, 1).toolTip())
        return out


class ExistsDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)

        text = opt.text
        opt.text = ""  # stop the default text paint (which uses QSS color)

        # Draw the normal cell (background, focus, selection, etc.)
        opt.widget.style().drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)

        # Now draw our text with our own pen color
        painter.save()
        if not (opt.state & QStyle.State_Selected):
            flag = index.data(Qt.UserRole)
            if flag is True:
                painter.setPen(QColor(120, 199, 143))
            elif flag is False:
                painter.setPen(QColor(255, 153, 160))
        painter.drawText(opt.rect.adjusted(6, 0, -6, 0),
                         Qt.AlignVCenter | Qt.AlignLeft, text)
        painter.restore()


class MainWindow(QMainWindow):
    class Columns(IntEnum):
        IDX = 0
        EXEC = 1
        PREF = 2
        EXISTS = 3

    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPU Preferences Manager")
        self.warning_bg = QColor(255, 200, 200)

        self.settings = QSettings("MyCompany", "GPU Preferences Manager")
        self.label_power = self.settings.value("label_power", "Power Saving")
        self.label_perf = self.settings.value("label_perf", "High Performance")
        self.default_gpu_high_perf: bool = self.settings.value(
            "default_gpu_high_perf", True, type=bool
        )

        self._apply_modern_style()
        self._init_window_size()
        self._build_ui()
        self._post_build()
        self._load_entries()

    def _apply_modern_style(self):
        app = QApplication.instance()
        if not app:
            return
        qss_path = Path(__file__).resolve().parent / "assets" / "theme.qss"
        try:
            with open(qss_path, "r", encoding="utf-8") as f:
                app.setStyleSheet(f.read())
        except Exception:
            p = QPalette()
            p.setColor(QPalette.Window, QColor(17, 19, 21))
            p.setColor(QPalette.WindowText, Qt.white)
            p.setColor(QPalette.Base, QColor(12, 14, 16))
            p.setColor(QPalette.AlternateBase, QColor(17, 19, 21))
            p.setColor(QPalette.ToolTipBase, Qt.white)
            p.setColor(QPalette.ToolTipText, Qt.white)
            p.setColor(QPalette.Text, Qt.white)
            p.setColor(QPalette.Button, QColor(17, 19, 21))
            p.setColor(QPalette.ButtonText, Qt.white)
            p.setColor(QPalette.Highlight, QColor(37, 99, 235))
            p.setColor(QPalette.HighlightedText, Qt.white)
            app.setPalette(p)

    def _init_window_size(self):
        size = self.settings.value("window_size")
        if isinstance(size, QSize):
            self.resize(size)
        else:
            try:
                if size:
                    w, h = map(int, size)
                    self.resize(QSize(w, h))
                else:
                    self.resize(980, 600)
            except Exception:
                self.resize(980, 600)

        screen = QGuiApplication.primaryScreen().availableGeometry()
        frame = self.frameGeometry()
        frame.moveCenter(screen.center())
        self.move(frame.topLeft())

    def _build_ui(self):
        self.statusBar()

        tb = self.addToolBar("Actions")
        tb.setMovable(False)

        add_menu = QMenu("Add", self)
        add_menu.addAction("Add Files…", self._on_add_files)
        add_menu.addAction("Add Folder…", self._on_add_folder)
        add_menu.addAction("From Running Processes…", self._on_add_running)

        add_btn = QToolButton(self)
        add_btn.setText("Add")
        add_btn.setMenu(add_menu)
        add_btn.setPopupMode(QToolButton.InstantPopup)
        add_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        set_role(add_btn, "primary")
        tb.addWidget(add_btn)

        self.act_remove = QAction(
            "Remove Selected", self, enabled=False
        )
        self.act_remove.triggered.connect(self._on_remove_selected)

        remove_btn = QToolButton(self)
        remove_btn.setDefaultAction(self.act_remove)
        remove_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        set_role(remove_btn, "danger")
        tb.addWidget(remove_btn)

        tb.addSeparator()

        self._seg = SegmentedControl([self.label_power, self.label_perf], self)
        self._seg.set_checked(None)
        self._seg.on_change(0, lambda: self._on_change_selected(False))
        self._seg.on_change(1, lambda: self._on_change_selected(True))
        tb.addWidget(self._seg)

        tb.addSeparator()

        # --- Check button: full-button menu, static label + dynamic menu ---
        self.act_check_selected = QAction("Check Selected", self)
        self.act_check_selected.triggered.connect(
            lambda: self._on_check_existence(True)
        )

        self.act_check_all = QAction("Check All", self)
        self.act_check_all.triggered.connect(lambda: self._on_check_existence(False))

        self.check_menu = QMenu(self)

        self.check_btn = QToolButton(self)
        self.check_btn.setText("Check Path")  # static label
        self.check_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.check_btn.setMenu(self.check_menu)
        self.check_btn.setPopupMode(QToolButton.InstantPopup)  # no split
        tb.addWidget(self.check_btn)

        act_refresh = QAction("Refresh", self)
        act_refresh.triggered.connect(self._load_entries)
        refresh_btn = QToolButton(self)
        refresh_btn.setDefaultAction(act_refresh)
        refresh_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        set_role(refresh_btn, "ghost")
        tb.addWidget(refresh_btn)

        tb.addSeparator()

        opt_menu = QMenu("Options", self)

        edit_labels = QAction("Customize GPU Labels…", self)
        edit_labels.triggered.connect(self._on_customize_labels)
        opt_menu.addAction(edit_labels)
        opt_menu.addSeparator()

        default_group = QActionGroup(self)
        self.act_default_power = QAction(self.label_power, self, checkable=True)
        self.act_default_perf = QAction(self.label_perf, self, checkable=True)
        default_group.addAction(self.act_default_power)
        default_group.addAction(self.act_default_perf)
        self.act_default_perf.setChecked(self.default_gpu_high_perf)
        self.act_default_power.setChecked(not self.default_gpu_high_perf)
        self.act_default_power.triggered.connect(lambda: self._set_default_gpu(False))
        self.act_default_perf.triggered.connect(lambda: self._set_default_gpu(True))

        default_gpu_menu = QMenu("Default GPU Preference", self)
        default_gpu_menu.addAction(self.act_default_power)
        default_gpu_menu.addAction(self.act_default_perf)
        opt_menu.addMenu(default_gpu_menu)

        opt_menu.addAction("Backup…", self._on_backup_config)
        opt_menu.addAction("Restore…", self._on_restore_config)

        opt_btn = QToolButton(self)
        opt_btn.setText("Options")
        opt_btn.setMenu(opt_menu)
        opt_btn.setPopupMode(QToolButton.InstantPopup)
        set_role(opt_btn, "ghost")
        tb.addWidget(opt_btn)

        spacer = QWidget(self)
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        tb.addWidget(spacer)

        self.quick_filter = QLineEdit(self, placeholderText="Filter executables…")
        self.quick_filter.textChanged.connect(self._apply_quick_filter)
        self.quick_filter.setMaximumWidth(280)
        tb.addWidget(self.quick_filter)

        self.table = QTableView(self)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.table.verticalHeader().hide()

        self.model = QStandardItemModel(0, 4, self)
        self.model.setHeaderData(self.Columns.IDX, Qt.Horizontal, "")
        self.model.setHeaderData(self.Columns.EXEC, Qt.Horizontal, "Executable")
        self.model.setHeaderData(self.Columns.PREF, Qt.Horizontal, "GPU Preference")
        self.model.setHeaderData(self.Columns.EXISTS, Qt.Horizontal, "Exists")

        self.table.setModel(self.model)
        self.table.setItemDelegateForColumn(self.Columns.EXEC, PathDelegate(self.table))
        self.table.setItemDelegateForColumn(self.Columns.PREF, PrefDelegate(self.table))
        self.table.setItemDelegateForColumn(
            self.Columns.EXISTS, ExistsDelegate(self.table)
        )

        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        hdr.setSectionResizeMode(self.Columns.IDX, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(self.Columns.EXEC, QHeaderView.Interactive)
        hdr.setSectionResizeMode(self.Columns.PREF, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(self.Columns.EXISTS, QHeaderView.ResizeToContents)
        self.table.setColumnWidth(self.Columns.EXISTS, 85)
        hdr.setStretchLastSection(False)
        hdr.setContextMenuPolicy(Qt.CustomContextMenu)
        hdr.customContextMenuRequested.connect(self._show_header_menu)

        central = QWidget(self)
        lay = QVBoxLayout(central)
        lay.setContentsMargins(12, 8, 12, 12)
        lay.addWidget(self.table)
        self.setCentralWidget(central)

    def _post_build(self):
        hdr = self.table.horizontalHeader()
        self._default_col_sizes = [hdr.sectionSize(i) for i in range(hdr.count())]

        for i in range(hdr.count()):
            w = self.settings.value(f"column_width_{i}")
            if w:
                try:
                    hdr.resizeSection(i, int(w))
                except Exception:
                    pass

        sel_model = self.table.selectionModel()
        sel_model.selectionChanged.connect(self._update_actions_state)
        self._update_actions_state()

    def _show_header_menu(self, pos):
        menu = QMenu(self)
        reset = QAction("Reset Column Widths", self)
        reset.triggered.connect(self._reset_column_widths)
        menu.addAction(reset)
        menu.exec(self.table.horizontalHeader().mapToGlobal(pos))

    def _reset_column_widths(self):
        hdr = self.table.horizontalHeader()
        for i, w in enumerate(self._default_col_sizes):
            hdr.resizeSection(i, w)
            self.settings.remove(f"column_width_{i}")

    def _load_entries(self):
        self.model.removeRows(0, self.model.rowCount())
        entries = RegistryManager.read_all()
        width = max(2, len(str(len(entries))))

        for i, e in enumerate(entries, start=1):
            idx_item = QStandardItem(f"{i:0{width}d}")

            exe_item = QStandardItem(e.exe)
            exe_item.setToolTip(e.exe)

            pref_item = QStandardItem(
                self.label_perf if e.pref == Pref.PERF else self.label_power
            )
            exists_item = QStandardItem()
            exists_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            self.model.appendRow([idx_item, exe_item, pref_item, exists_item])
            self._update_row_existence(self.model.rowCount() - 1)

        self.table.resizeColumnToContents(self.Columns.IDX)

    def _apply_quick_filter(self, text: str):
        text = text.strip().lower()
        for r in range(self.model.rowCount()):
            exec_txt = self.model.item(r, self.Columns.EXEC).text().lower()
            pref_txt = self.model.item(r, self.Columns.PREF).text().lower()
            match = (text in exec_txt) or (text in pref_txt)
            self.table.setRowHidden(r, not match if text else False)

    def _update_row_existence(self, row: int) -> None:
        exe = self.model.item(row, self.Columns.EXEC).text()
        exists_text = ""
        exists_flag: Optional[bool] = None
        if is_exe(exe):
            exists_flag = Path(exe).exists()
            exists_text = "●  Yes" if exists_flag else "●  No"

        item = self.model.item(row, self.Columns.EXISTS)
        item.setText(exists_text)
        # Store boolean for the delegate; None means "unknown/blank"
        item.setData(exists_flag, Qt.UserRole)

        # Clear previous manual roles for the whole row
        for c in range(self.model.columnCount()):
            it = self.model.item(row, c)
            it.setData(None, Qt.BackgroundRole)
            it.setData(None, Qt.ForegroundRole)

        # Keep the missing-path background highlight on the EXEC cell
        if exists_flag is False:
            warn_bg = QColor(44, 7, 7)
            self.model.item(row, self.Columns.EXEC).setBackground(warn_bg)

    def _existing_set(self) -> set[str]:
        return {e.exe for e in RegistryManager.read_all()}

    def _add_paths(self, paths: Iterable[str]):
        existing = self._existing_set()
        default_pref = Pref.PERF if self.default_gpu_high_perf else Pref.POWER
        added, skipped = 0, 0
        for p in paths:
            absp = normalize_path(Path(p).resolve())
            if absp in existing:
                skipped += 1
                continue
            RegistryManager.set_pref(absp, default_pref)
            added += 1

        self._load_entries()
        if added or skipped:
            msg = []
            if added:
                msg.append(f"Added {added}.")
            if skipped:
                msg.append(f"Ignored {skipped} duplicate(s).")
            QMessageBox.information(self, "Results", " ".join(msg))

    def _on_add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select EXEs", str(Path.home()), "Executables (*.exe)"
        )
        if paths:
            self._add_paths(paths)

    def _on_add_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder", str(Path.home())
        )
        if not folder:
            return
        paths: List[str] = []
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".exe"):
                    paths.append(str(Path(root) / f))
        self._add_paths(paths)

    def _on_add_running(self):
        dlg = RunningProcessDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self._add_paths(dlg.selected_paths())

    def _on_remove_selected(self):
        rows = [i.row() for i in self.table.selectionModel().selectedRows()]
        paths = [self.model.item(r, self.Columns.EXEC).text() for r in rows]
        for p in paths:
            RegistryManager.delete(p)
        self._load_entries()

    def _on_change_selected(self, high_perf: bool):
        rows = [i.row() for i in self.table.selectionModel().selectedRows()]
        if not rows:
            return
        for r in rows:
            exe = self.model.item(r, self.Columns.EXEC).text()
            RegistryManager.set_pref(exe, Pref.PERF if high_perf else Pref.POWER)
            self.model.item(r, self.Columns.PREF).setText(
                self.label_perf if high_perf else self.label_power
            )
        self._load_entries()
        self._refresh_segmented_from_selection()

    def _on_check_existence(self, selected_only: bool):
        if selected_only:
            idxs = self.table.selectionModel().selectedRows()
            rows = [i.row() for i in idxs] if idxs else []
            if not rows:
                return
        else:
            rows = list(range(self.model.rowCount()))

        for r in rows:
            self._update_row_existence(r)

    def _on_backup_config(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Backup Config As…", str(Path.home()), "JSON Files (*.json)"
        )
        if not path:
            return
        RegistryManager.backup(path)
        QMessageBox.information(self, "Backup", f"Saved to:\n{path}")

    def _on_restore_config(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Backup…", str(Path.home()), "JSON Files (*.json)"
        )
        if not path:
            return
        RegistryManager.restore(path)
        self._load_entries()
        QMessageBox.information(self, "Restore", "Configuration restored.")
        self._refresh_segmented_from_selection()

    def _set_default_gpu(self, high_performance: bool):
        self.default_gpu_high_perf = high_performance
        self.settings.setValue("default_gpu_high_perf", high_performance)
        self.act_default_perf.setChecked(high_performance)
        self.act_default_power.setChecked(
            high_performance if False else not high_performance
        )
        self._seg.set_checked(None)
        QMessageBox.information(
            self,
            "Default GPU Preference",
            f"New default set to “{self.label_perf if high_performance else self.label_power}.”",
        )

    def _on_customize_labels(self):
        p1, ok1 = QInputDialog.getText(
            self,
            "Label for Power-Saving GPU",
            "Name for the power-saving profile:",
            text=self.label_power,
        )
        if ok1 and p1.strip():
            self.label_power = p1.strip()
            self.settings.setValue("label_power", self.label_power)

        p2, ok2 = QInputDialog.getText(
            self,
            "Label for High-Performance GPU",
            "Name for the high-performance profile:",
            text=self.label_perf,
        )
        if ok2 and p2.strip():
            self.label_perf = p2.strip()
            self.settings.setValue("label_perf", self.label_perf)

        self._load_entries()

        self.table.setItemDelegateForColumn(self.Columns.PREF, PrefDelegate(self.table))
        self.act_default_power.setText(self.label_power)
        self.act_default_perf.setText(self.label_perf)
        self._seg.set_texts([self.label_power, self.label_perf])

        self._refresh_segmented_from_selection()

    def _refresh_segmented_from_selection(self):
        """Reflect the selected row(s) preference in the segmented control;
        if 0 rows or >1 rows are selected, clear both."""
        idxs = self.table.selectionModel().selectedRows()
        if len(idxs) != 1:
            self._seg.set_checked(None)
            return
        row = idxs[0].row()
        pref_text = self.model.item(row, self.Columns.PREF).text()
        if pref_text == self.label_perf:
            self._seg.set_checked(1)
        else:
            self._seg.set_checked(0)

    def _update_actions_state(self, *_):
        has_sel = bool(self.table.selectionModel().selectedRows())
        sel_rows = len(self.table.selectionModel().selectedRows())  # noqa F841
        self.act_remove.setEnabled(has_sel)
        for b in self._seg.buttons:
            b.setEnabled(has_sel)

        # Static check button label; no dynamic text.

        # Dynamic menu ordering + default
        self.check_menu.clear()
        if has_sel:
            self.check_menu.addAction(self.act_check_selected)
            self.check_menu.addAction(self.act_check_all)
            self.check_menu.setDefaultAction(self.act_check_selected)
        else:
            self.check_menu.addAction(self.act_check_all)
            self.check_menu.addAction(self.act_check_selected)
            self.check_menu.setDefaultAction(self.act_check_all)

        self._refresh_segmented_from_selection()

    def closeEvent(self, event):
        self.settings.setValue("window_size", self.size())
        hdr = self.table.horizontalHeader()
        for i in range(hdr.count()):
            self.settings.setValue(f"column_width_{i}", hdr.sectionSize(i))
        super().closeEvent(event)


def set_taskbar_icon(hwnd: int, icon_path: str):
    IMAGE_ICON = 1
    LR_LOADFROMFILE = 0x00000010
    LR_DEFAULTSIZE = 0x00000040
    flags = LR_LOADFROMFILE | LR_DEFAULTSIZE
    hicon = ctypes.windll.user32.LoadImageW(None, icon_path, IMAGE_ICON, 0, 0, flags)
    if hicon:
        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon)
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon)


def main():
    current_dir = Path(__file__).resolve().parent
    assets_dir = current_dir / "assets"
    assets_dir.mkdir(exist_ok=True)
    icon_path = str(assets_dir / "icon.ico")

    app = QApplication(sys.argv)
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    win = MainWindow()
    win.show()

    if os.path.exists(icon_path):
        set_taskbar_icon(int(win.winId()), icon_path)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
