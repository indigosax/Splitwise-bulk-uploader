@echo off
echo ============================================
echo  Splitwise CSV Importer — Build EXE
echo ============================================
echo.

:: Install dependencies
echo [1/3] Installing dependencies...
pip install requests pyinstaller --quiet
if errorlevel 1 (
    echo ERROR: pip install failed. Make sure Python 3.10+ is on PATH.
    pause
    exit /b 1
)

:: Build the executable
echo [2/3] Building executable...
pyinstaller --onefile ^
            --windowed ^
            --name "SplitwiseImporter" ^
            --add-data "sample_expenses.csv;." ^
            splitwise_importer.py

if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

:: Done
echo [3/3] Done!
echo.
echo  Your executable is at:
echo    dist\SplitwiseImporter.exe
echo.
echo  Copy the .exe anywhere — no Python required on the target machine.
echo.
pause
