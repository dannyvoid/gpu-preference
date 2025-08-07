#!/usr/bin/env python3
import os
import sys
import json
import ctypes
import winreg
import psutil

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QTableView,
    QFileDialog,
    QMessageBox,
    QAbstractItemView,
    QHeaderView,
    QToolButton,
    QMenu,
    QLineEdit,
    QComboBox,
    QStyledItemDelegate,
    QDialog,
    QDialogButtonBox,
    QInputDialog,
)
from PySide6.QtGui import (
    QAction,
    QStandardItemModel,
    QStandardItem,
    QPalette,
    QColor,
    QIcon,
    QGuiApplication,
)
from PySide6.QtCore import Qt, QSize, QSortFilterProxyModel, QSettings

REG_PATH = r"Software\Microsoft\DirectX\UserGpuPreferences"


def normalize_path(p: str) -> str:
    path = os.path.normpath(p)
    if len(path) >= 2 and path[1] == ":":
        path = path[0].upper() + path[1:]
    return path


def get_registry_entries():
    entries = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
            i = 0
            while True:
                name, val, _ = winreg.EnumValue(key, i)
                name = normalize_path(name)
                code = 1 if "GpuPreference=1" in val else 2
                entries.append((name, code))
                i += 1
    except OSError:
        pass
    return entries


def set_registry_entry(exe_path, high_performance: bool):
    exe_norm = normalize_path(exe_path)
    pref = 2 if high_performance else 1
    with winreg.CreateKeyEx(
        winreg.HKEY_CURRENT_USER,
        REG_PATH,
        0,
        winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY,
    ) as key:
        winreg.SetValueEx(key, exe_norm, 0, winreg.REG_SZ, f"GpuPreference={pref};")


def delete_registry_entry(exe_path):
    exe_norm = normalize_path(exe_path)
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            REG_PATH,
            0,
            winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY,
        ) as key:
            winreg.DeleteValue(key, exe_norm)
    except FileNotFoundError:
        pass


def contrasting_text_color(bg: QColor) -> QColor:
    colors = (0.299 * bg.red() + 0.587 * bg.green() + 0.114 * bg.blue()) / 255
    return QColor(0, 0, 0) if colors > 0.5 else QColor(255, 255, 255)


class PathDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return QLineEdit(parent)

    def setEditorData(self, editor, index):
        editor.setText(index.model().data(index, Qt.EditRole))

    def setModelData(self, editor, model, index):
        raw = editor.text()
        new = normalize_path(raw)
        old = model.item(index.row(), 1).text()

        if new != old:
            code = (
                1
                if model.item(index.row(), 2).text() == model.parent().label_power
                else 2
            )
            delete_registry_entry(old)
            set_registry_entry(new, high_performance=(code == 2))

        model.setData(index, new, Qt.EditRole)

        row = index.row()
        exists_item = model.item(row, 3)
        if os.path.isabs(new) and new.lower().endswith(".exe"):
            exists = os.path.exists(new)
            exists_item.setText("Yes" if exists else "No")
        else:
            exists = None
            exists_item.setText("")

        mw = model.parent()

        if exists is False:
            fg = contrasting_text_color(mw.warning_bg)
            for col in range(model.columnCount()):
                item = model.item(row, col)
                item.setBackground(mw.warning_bg)
                item.setForeground(fg)
        else:
            for col in range(model.columnCount()):
                item = model.item(row, col)
                bg = item.background()
                if bg.style() != Qt.NoBrush and bg.color() == mw.warning_bg:
                    item.setData(None, Qt.BackgroundRole)
                    item.setData(None, Qt.ForegroundRole)


class PrefDelegate(QStyledItemDelegate):
    def __init__(self, parent, label_power, label_perf):
        super().__init__(parent)
        self.label_power = label_power
        self.label_perf = label_perf

    def createEditor(self, parent, option, index):
        cb = QComboBox(parent)
        cb.addItems([self.label_power, self.label_perf])
        return cb

    def setEditorData(self, editor, index):
        current = index.model().data(index, Qt.EditRole)
        editor.setCurrentIndex(0 if self.label_power in current else 1)

    def setModelData(self, editor, model, index):
        text = editor.currentText()
        exe = model.item(index.row(), 1).text()
        set_registry_entry(exe, high_performance=(editor.currentIndex() == 1))
        model.setData(index, text, Qt.EditRole)


class RunningProcessDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add from Running Processes")
        self.resize(600, 400)

        layout = QVBoxLayout(self)
        self.search = QLineEdit(self)
        self.search.setPlaceholderText("Search executables…")
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
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectRows
        )  # <-- fixed indent
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().hide()

        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Interactive)
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
        seen = set()
        for proc in psutil.process_iter(["exe", "pid"]):
            path = proc.info["exe"]
            if not path:
                continue
            path = normalize_path(path)
            if not path.lower().endswith(".exe"):
                continue
            if path in seen:
                continue
            seen.add(path)
            chk = QStandardItem()
            chk.setCheckable(True)
            chk.setCheckState(Qt.Unchecked)
            exe_item = QStandardItem(os.path.basename(path))
            exe_item.setToolTip(path)
            pid_item = QStandardItem(str(proc.info["pid"]))
            self.src_model.appendRow([chk, exe_item, pid_item])

    def selected_paths(self):
        paths = []
        for row in range(self.src_model.rowCount()):
            if self.src_model.item(row, 0).checkState() == Qt.Checked:
                paths.append(self.src_model.item(row, 1).toolTip())
        return paths


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPU Preferences Manager")

        self.settings = QSettings("MyCompany", "GPU Preferences Manager")

        self.label_power = self.settings.value("label_power", "Power Saving")
        self.label_perf = self.settings.value("label_perf", "High Performance")

        self.default_gpu_high_perf = self.settings.value(
            "default_gpu_high_perf", True, type=bool
        )

        size = self.settings.value("window_size")
        if isinstance(size, QSize):
            self.resize(size)
        elif size:
            try:
                w, h = map(int, size)
                self.resize(QSize(w, h))
            except (ValueError, TypeError):
                self.resize(QSize(950, 550))
        else:
            self.resize(QSize(950, 550))

        self.warning_bg = QColor(255, 200, 200)

        self._setup_ui()
        self.center_window()

        hdr = self.table.horizontalHeader()
        self.default_column_sizes = [hdr.sectionSize(i) for i in range(hdr.count())]

        self.load_entries()
        for i in range(hdr.count()):
            w = self.settings.value(f"column_width_{i}")
            if w:
                try:
                    hdr.resizeSection(i, int(w))
                except (ValueError, TypeError):
                    pass

        self.table.selectionModel().selectionChanged.connect(self._update_actions_state)

    def center_window(self):
        screen = QGuiApplication.primaryScreen().availableGeometry()
        frame = self.frameGeometry()
        frame.moveCenter(screen.center())
        self.move(frame.topLeft())

    def _apply_dark_palette(self):
        p = QPalette()
        p.setColor(QPalette.Window, QColor(53, 53, 53))
        p.setColor(QPalette.WindowText, Qt.white)
        p.setColor(QPalette.Base, QColor(35, 35, 35))
        p.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        p.setColor(QPalette.ToolTipBase, Qt.white)
        p.setColor(QPalette.ToolTipText, Qt.white)
        p.setColor(QPalette.Text, Qt.white)
        p.setColor(QPalette.Button, QColor(53, 53, 53))
        p.setColor(QPalette.ButtonText, Qt.white)
        p.setColor(QPalette.BrightText, Qt.red)
        p.setColor(QPalette.Highlight, QColor(0, 120, 215))
        p.setColor(QPalette.HighlightedText, Qt.white)
        QApplication.setPalette(p)

    def _setup_ui(self):
        QApplication.setStyle("Fusion")
        self._apply_dark_palette()

        tb = self.addToolBar("MainToolbar")
        tb.setMovable(False)

        add_menu = QMenu("Add", self)
        add_menu.addAction("Add Files…", self.on_add_files)
        add_menu.addAction("Add Folder…", self.on_add_folder)
        add_menu.addAction("From Running Processes…", self.on_add_running)
        add_btn = QToolButton(self)
        add_btn.setText("Add")
        add_btn.setMenu(add_menu)
        add_btn.setPopupMode(QToolButton.MenuButtonPopup)
        add_btn.setIcon(QIcon.fromTheme("list-add"))
        tb.addWidget(add_btn)
        tb.addSeparator()

        self.act_remove = QAction(
            QIcon.fromTheme("edit-delete"), "Remove Selected", self
        )
        self.act_remove.setEnabled(False)
        self.act_remove.triggered.connect(self.on_remove_selected)
        tb.addAction(self.act_remove)

        pref_menu = QMenu("Set GPU Preference", self)
        pref_menu.addAction(self.label_power, lambda: self.on_change_selected(False))
        pref_menu.addAction(self.label_perf, lambda: self.on_change_selected(True))
        self.act_pref = QToolButton(self)
        self.act_pref.setText("GPU Pref")
        self.act_pref.setMenu(pref_menu)
        self.act_pref.setPopupMode(QToolButton.MenuButtonPopup)
        self.act_pref.setIcon(QIcon.fromTheme("preferences-system"))
        self.act_pref.setEnabled(False)
        tb.addWidget(self.act_pref)
        tb.addSeparator()

        self.check_btn = QToolButton(self)
        self.check_btn.setIcon(QIcon.fromTheme("system-search"))
        self.check_btn.setToolTip("Check selected rows (or all if none selected)")
        self.check_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.check_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.check_btn.clicked.connect(
            lambda: self.on_check_existence(selected_only=True)
        )
        check_menu = QMenu(self)
        check_menu.addAction("Check Selected", lambda: self.on_check_existence(True))
        check_menu.addAction("Check All", lambda: self.on_check_existence(False))
        self.check_btn.setMenu(check_menu)
        tb.addWidget(self.check_btn)
        tb.addSeparator()

        act_refresh = QAction(QIcon.fromTheme("view-refresh"), "Refresh", self)
        act_refresh.triggered.connect(self.load_entries)
        tb.addAction(act_refresh)
        tb.addSeparator()

        opt_menu = QMenu("Options", self)
        edit_labels = QAction("Customize GPU Labels…", self)
        edit_labels.triggered.connect(self.on_customize_labels)
        opt_menu.addAction(edit_labels)
        opt_menu.addSeparator()
        default_gpu_menu = QMenu("Default GPU Preference", self)
        self.act_default_power = QAction(self.label_power, self, checkable=True)
        self.act_default_perf = QAction(self.label_perf, self, checkable=True)
        self.act_default_power.triggered.connect(lambda: self.set_default_gpu(False))
        self.act_default_perf.triggered.connect(lambda: self.set_default_gpu(True))
        default_gpu_menu.addAction(self.act_default_power)
        default_gpu_menu.addAction(self.act_default_perf)
        self.act_default_perf.setChecked(self.default_gpu_high_perf)
        self.act_default_power.setChecked(not self.default_gpu_high_perf)
        opt_menu.addMenu(default_gpu_menu)
        opt_menu.addAction("Backup…", self.on_backup_config)
        opt_menu.addAction("Restore…", self.on_restore_config)
        opt_btn = QToolButton(self)
        opt_btn.setText("Options")
        opt_btn.setMenu(opt_menu)
        opt_btn.setPopupMode(QToolButton.InstantPopup)
        opt_btn.setIcon(QIcon.fromTheme("applications-system"))
        tb.addWidget(opt_btn)

        self.table = QTableView()
        self.table.setModel(QStandardItemModel(0, 4, self))
        self.model = self.table.model()
        self.model.setHeaderData(0, Qt.Horizontal, "Index")
        self.model.setHeaderData(1, Qt.Horizontal, "Executable")
        self.model.setHeaderData(2, Qt.Horizontal, "GPU Preference")
        self.model.setHeaderData(3, Qt.Horizontal, "Exists")
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSortingEnabled(True)
        self.table.verticalHeader().hide()
        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Interactive)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.setColumnWidth(1, 350)
        self.table.setColumnWidth(2, 200)
        self.table.setColumnWidth(3, 70)
        hdr.setStretchLastSection(False)
        hdr.setContextMenuPolicy(Qt.CustomContextMenu)
        hdr.customContextMenuRequested.connect(self._show_header_menu)
        self.table.setItemDelegateForColumn(1, PathDelegate(self.table))
        self.table.setItemDelegateForColumn(
            2, PrefDelegate(self.table, self.label_power, self.label_perf)
        )

        central = QWidget()
        self.setCentralWidget(central)
        vlay = QVBoxLayout(central)
        vlay.addWidget(self.table)

        self._update_actions_state()

    def _show_header_menu(self, pos):
        menu = QMenu(self)
        reset = QAction("Reset Column Widths", self)
        reset.triggered.connect(self._reset_column_widths)
        menu.addAction(reset)
        menu.exec(self.table.horizontalHeader().mapToGlobal(pos))

    def _reset_column_widths(self):
        hdr = self.table.horizontalHeader()
        for i, w in enumerate(self.default_column_sizes):
            hdr.resizeSection(i, w)
            self.settings.remove(f"column_width_{i}")

    def _update_actions_state(self, *_):
        sel = self.table.selectionModel().selectedRows()
        has = bool(sel)
        self.act_remove.setEnabled(has)
        self.act_pref.setEnabled(has)
        if has:
            self.check_btn.setText(f"Check {len(sel)} Selected")
        else:
            self.check_btn.setText("Check All")

    def load_entries(self):
        self.model.removeRows(0, self.model.rowCount())
        entries = get_registry_entries()
        width = max(2, len(str(len(entries))))
        for i, (exe, code) in enumerate(entries, start=1):
            idx = QStandardItem(f"{i:0{width}d}")
            exe_item = QStandardItem(exe)
            exe_item.setToolTip(exe)
            pref_item = QStandardItem(
                self.label_power if code == 1 else self.label_perf
            )

            if os.path.isabs(exe) and exe.lower().endswith(".exe"):
                exists = os.path.exists(exe)
                text = "Yes" if exists else "No"
            else:
                exists = None
                text = ""
            exists_item = QStandardItem(text)
            exists_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            self.model.appendRow([idx, exe_item, pref_item, exists_item])
            self.table.resizeColumnToContents(0)

            if exists is False:
                fg = contrasting_text_color(self.warning_bg)
                row = self.model.rowCount() - 1
                for col in range(self.model.columnCount()):
                    item = self.model.item(row, col)
                    item.setBackground(self.warning_bg)
                    item.setForeground(fg)

    def on_add_files(self):
        existing = {p for p, _ in get_registry_entries()}
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select EXEs", os.path.expanduser("~"), "Executables (*.exe)"
        )
        skipped = []
        default = self.default_gpu_high_perf
        for p in paths:
            absp = normalize_path(os.path.abspath(p))
            if absp in existing:
                skipped.append(absp)
            else:
                set_registry_entry(absp, high_performance=default)
        self.load_entries()
        if skipped:
            QMessageBox.information(
                self,
                "Skipped Duplicates",
                f"Ignored {len(skipped)} already-added executables.",
            )

    def on_add_folder(self):
        existing = {p for p, _ in get_registry_entries()}
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder", os.path.expanduser("~")
        )
        skipped = []
        if folder:
            default = self.default_gpu_high_perf
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith(".exe"):
                        absp = normalize_path(os.path.join(root, f))
                        if absp in existing:
                            skipped.append(absp)
                        else:
                            set_registry_entry(absp, high_performance=default)
        self.load_entries()
        if skipped:
            QMessageBox.information(
                self,
                "Skipped Duplicates",
                f"Ignored {len(skipped)} already-added executables.",
            )

    def on_add_running(self):
        existing = {p for p, _ in get_registry_entries()}
        dlg = RunningProcessDialog(self)
        if dlg.exec() == QDialog.Accepted:
            skipped, added = [], []
            default = self.default_gpu_high_perf
            for path in dlg.selected_paths():
                if path in existing:
                    skipped.append(path)
                else:
                    set_registry_entry(path, high_performance=default)
                    added.append(path)
            self.load_entries()
            msg = []
            if added:
                msg.append(f"Added {len(added)} process(es).")
            if skipped:
                msg.append(f"Ignored {len(skipped)} duplicate(s).")
            QMessageBox.information(self, "Results", "\n".join(msg))

    def on_remove_selected(self):
        for idx in self.table.selectionModel().selectedRows():
            delete_registry_entry(self.model.item(idx.row(), 1).text())
        self.load_entries()

    def on_change_selected(self, high_perf: bool):
        for idx in self.table.selectionModel().selectedRows():
            set_registry_entry(
                self.model.item(idx.row(), 1).text(), high_performance=high_perf
            )
        self.load_entries()

    def on_check_existence(self, selected_only=True):
        if selected_only:
            idxs = self.table.selectionModel().selectedRows()
            rows = [i.row() for i in idxs] if idxs else []
        else:
            rows = range(self.model.rowCount())

        for row in rows:
            exe = self.model.item(row, 1).text()
            if os.path.isabs(exe) and exe.lower().endswith(".exe"):
                exists = os.path.exists(exe)
                text = "Yes" if exists else "No"
            else:
                exists = None
                text = ""
            self.model.item(row, 3).setText(text)

            if exists is False:
                fg = contrasting_text_color(self.warning_bg)
                for col in range(self.model.columnCount()):
                    item = self.model.item(row, col)
                    item.setBackground(self.warning_bg)
                    item.setForeground(fg)
            else:
                for col in range(self.model.columnCount()):
                    item = self.model.item(row, col)
                    bg = item.background()
                    if bg.style() != Qt.NoBrush and bg.color() == self.warning_bg:
                        item.setData(None, Qt.BackgroundRole)
                        item.setData(None, Qt.ForegroundRole)

    def on_backup_config(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Backup Config As…", os.path.expanduser("~"), "JSON Files (*.json)"
        )
        if not path:
            return
        data = {exe: code for exe, code in get_registry_entries()}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        QMessageBox.information(self, "Backup", f"Saved to:\n{path}")

    def on_restore_config(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Backup…", os.path.expanduser("~"), "JSON Files (*.json)"
        )
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for exe, _ in get_registry_entries():
            delete_registry_entry(exe)
        for exe, code in data.items():
            set_registry_entry(exe, high_performance=(code == 2))
        self.load_entries()
        QMessageBox.information(self, "Restore", "Configuration restored.")

    def set_default_gpu(self, high_performance: bool):
        self.default_gpu_high_perf = high_performance
        self.settings.setValue("default_gpu_high_perf", high_performance)
        self.act_default_perf.setChecked(high_performance)
        self.act_default_power.setChecked(not high_performance)
        QMessageBox.information(
            self,
            "Default GPU Preference",
            f"New default set to “{self.label_perf if high_performance else self.label_power}.”",
        )

    def on_customize_labels(self):
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

        self.load_entries()
        self.table.setItemDelegateForColumn(
            2, PrefDelegate(self.table, self.label_power, self.label_perf)
        )
        self.act_default_power.setText(self.label_power)
        self.act_default_perf.setText(self.label_perf)

    def closeEvent(self, event):
        self.settings.setValue("window_size", self.size())
        hdr = self.table.horizontalHeader()
        for i in range(hdr.count()):
            self.settings.setValue(f"column_width_{i}", hdr.sectionSize(i))
        super().closeEvent(event)


def set_taskbar_icon(hwnd, icon_path):
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
    current_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(current_dir, "assets", "icon.ico")

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(icon_path))

    win = MainWindow()
    win.show()

    set_taskbar_icon(int(win.winId()), icon_path)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
