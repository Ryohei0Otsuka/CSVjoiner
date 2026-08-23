@echo off
setlocal
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --clean --onefile --noconsole --name CSVjoiner CSVjoiner.py

echo.
echo Build complete: dist\CSVjoiner.exe
pause
