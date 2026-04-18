@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%generate_akanecco_shift_gui.exe"

if not exist "%EXE_PATH%" (
    echo generate_akanecco_shift_gui.exe was not found.
    pause
    exit /b 1
)

start "" "%EXE_PATH%"
exit /b 0