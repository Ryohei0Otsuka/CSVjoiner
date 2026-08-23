@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo   CSVjoiner v2 - Windows EXE Build
echo ========================================
echo.

python --version || goto :error
python -m pip install -r requirements.txt || goto :error
python -m pip install -r requirements-dev.txt || goto :error

echo.
echo [1/2] Running tests...
python -m pytest -q || goto :error

echo.
echo [2/2] Building one-file EXE...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
python -m PyInstaller --noconfirm --clean CSVjoiner.spec || goto :error

if not exist "dist\CSVjoiner.exe" goto :error

echo.
echo ========================================
echo Build complete
echo   dist\CSVjoiner.exe
echo ========================================
echo.
pause
exit /b 0

:error
echo.
echo ========================================
echo Build failed. See the messages above.
echo ========================================
echo.
pause
exit /b 1
