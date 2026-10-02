@echo off
REM ==========================================================
REM  Build BM_DMCS_Tool.exe (single file, no console window)
REM  Usage: double-click this file, or run it in cmd.
REM  Requires Python 3.9+ ("py" launcher) installed on Windows.
REM ==========================================================
setlocal
cd /d "%~dp0"

set APP_NAME=BM_DMCS_Tool_v8_13_Beta
set SCRIPT=BM_DMCS_Tool_v8_13_Beta.py

echo [1/3] Installing build packages...
py -m pip install --upgrade pip
py -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error

echo [2/3] Building %APP_NAME%.exe ...
py -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "%APP_NAME%" ^
    --icon app_icon.ico ^
    "%SCRIPT%"
if errorlevel 1 goto :error

echo [3/3] Done.
echo.
echo   EXE : %~dp0dist\%APP_NAME%.exe
echo.
echo   Copy the EXE to any folder and run it. On first run it creates
echo   02_test_input, 03_reference, 04_output, 05_logs next to the EXE.
echo.
pause
exit /b 0

:error
echo.
echo Build FAILED. Check the messages above.
pause
exit /b 1
