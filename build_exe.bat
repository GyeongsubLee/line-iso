@echo off
rem ============================================================
rem  Line List / ISO QC Tool - release build (PyInstaller)
rem  Put this file, make_icon.py and line_iso_desktop_tool_V*.py
rem  in the same folder, then double-click.
rem  Result: release\LineIsoQC_<version>.zip
rem ============================================================
setlocal
cd /d "%~dp0"

rem --- pick the newest line_iso_desktop_tool_V*.py -------------
set "APP="
for /f "delims=" %%f in ('dir /b /on "line_iso_desktop_tool_V*.py" 2^>nul') do set "APP=%%f"
if "%APP%"=="" (
    echo [ERROR] line_iso_desktop_tool_V*.py not found in %CD%
    pause
    exit /b 1
)
set "VER=%APP:line_iso_desktop_tool_=%"
set "VER=%VER:.py=%"
set "NAME=LineIsoQC_%VER%"
echo [INFO] Source : %APP%
echo [INFO] Output : %NAME%

rem --- clean build environment (only the packages the tool needs) --
if not exist ".build_venv\Scripts\python.exe" (
    echo [INFO] Creating build virtual environment...
    py -m venv .build_venv || goto :fail
)
set "PY=.build_venv\Scripts\python.exe"
"%PY%" -m pip install --upgrade pip >nul
"%PY%" -m pip install pyinstaller PySide6 pandas openpyxl pywin32 pillow || goto :fail

rem --- exe icon from the icon embedded in the program -----------
"%PY%" make_icon.py "%APP%" app.ico || goto :fail

rem --- build: folder type (fast start, fewer antivirus issues) ---
"%PY%" -m PyInstaller --noconfirm --clean --windowed --onedir ^
    --name "%NAME%" --icon app.ico ^
    --exclude-module tkinter --exclude-module matplotlib --exclude-module IPython ^
    --exclude-module PySide6.QtWebEngineCore --exclude-module PySide6.Qt3DCore ^
    "%APP%" || goto :fail

rem --- shared settings: copy config\*.json except personal settings --
if exist "config" (
    mkdir "dist\%NAME%\config" 2>nul
    for %%j in (config\*.json) do (
        if /i not "%%~nxj"=="project_settings.json" if /i not "%%~nxj"=="project_mapping.json" copy /y "%%j" "dist\%NAME%\config\" >nul
    )
    echo [INFO] Copied shared config files ^(project_settings.json excluded^)
)

rem --- zip for distribution -------------------------------------
if not exist "release" mkdir "release"
if exist "release\%NAME%.zip" del "release\%NAME%.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\%NAME%' -DestinationPath 'release\%NAME%.zip'" || goto :fail

echo.
echo [DONE] dist\%NAME%\%NAME%.exe
echo [DONE] release\%NAME%.zip  ^<- send this file
pause
exit /b 0

:fail
echo.
echo [ERROR] Build failed. Check the messages above.
pause
exit /b 1
