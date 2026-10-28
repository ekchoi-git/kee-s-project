@echo off
rem Build Celly exe in a clean environment with the pinned Qt version.
rem Output: dist\Celly.exe
cd /d "%~dp0"
if not exist ".packaging-venv\Scripts\python.exe" python -m venv .packaging-venv
if errorlevel 1 exit /b 1
".packaging-venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 exit /b 1
".packaging-venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name Celly --icon "app\assets\celly-icon.ico" --add-data "app\assets;app\assets" --add-data "app\fonts;app\fonts" celly.py
if errorlevel 1 exit /b 1
