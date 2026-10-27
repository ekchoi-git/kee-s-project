@echo off
rem Build Celli exe. First time only: python -m pip install pyinstaller
rem Output: dist\Celli.exe (bundles app\fonts Pretendard; falls back to Malgun Gothic if missing)
cd /d "%~dp0"
python -m PyInstaller --noconfirm --onefile --windowed --name Celli --icon "app\assets\celli-icon.ico" --add-data "app\assets;app\assets" --add-data "app\fonts;app\fonts" celli.py
