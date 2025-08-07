@echo off

pushd %~dp0

pyinstaller --clean --onefile --windowed --name windowsGPU --add-data 'application/src/assets;assets' --icon 'application/src/assets/icon.ico' application/src/main.py
