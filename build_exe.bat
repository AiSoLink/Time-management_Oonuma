@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo Build EXE script
echo ========================================

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
    echo venv activated
)

python -c "import PyInstaller" 1>nul 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building...
echo.

pyinstaller --onefile --windowed --name "oonuma_uploader" --hidden-import openpyxl --hidden-import openpyxl.styles --clean "ダブルクリックボタン.py"
if errorlevel 1 (
    echo.
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed
echo ========================================
echo Output: dist\oonuma_uploader.exe
echo.
pause
