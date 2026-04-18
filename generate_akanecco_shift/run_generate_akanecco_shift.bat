@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%generate_akanecco_shift.exe"
set "RUNNER_PATH=%SCRIPT_DIR%run_with_utf8_log.ps1"
set "LOG_DIR=%SCRIPT_DIR%logs"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TIMESTAMP=%%I"
set "LOG_PATH=%LOG_DIR%\run_generate_akanecco_shift_%TIMESTAMP%.log"

if not exist "%EXE_PATH%" (
    echo generate_akanecco_shift.exe が見つかりません。
    echo 配置先を確認してください。
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[System.IO.File]::WriteAllLines('%LOG_PATH%', @('[%date% %time%] ERROR: generate_akanecco_shift.exe was not found.', '[%date% %time%] SCRIPT_DIR=%SCRIPT_DIR%'), [System.Text.UTF8Encoding]::new($false))"
    echo Log file: %LOG_PATH%
    call :maybe_pause
    exit /b 1
)

if not exist "%RUNNER_PATH%" (
    echo run_with_utf8_log.ps1 が見つかりません。
    echo 配置先を確認してください。
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[System.IO.File]::WriteAllLines('%LOG_PATH%', @('[%date% %time%] ERROR: run_with_utf8_log.ps1 was not found.', '[%date% %time%] SCRIPT_DIR=%SCRIPT_DIR%'), [System.Text.UTF8Encoding]::new($false))"
    echo Log file: %LOG_PATH%
    call :maybe_pause
    exit /b 1
)

if "%~1"=="" goto :help

if /i "%~x1"==".xls" goto :generate
if /i "%~x1"==".xlsx" goto :generate

call :run_with_log %*
call :maybe_pause
exit /b %EXIT_CODE%

:generate
echo Target file: %~1
call :run_with_log generate --target "%~1"
call :maybe_pause
exit /b %EXIT_CODE%

:help
echo Usage:
echo 1. Drag and drop a temp Excel file onto this bat to run generate.
echo 2. Run this bat with command-line args to pass them to the exe.
echo.
echo Examples:
echo   run_generate_akanecco_shift.bat "【統一書式】あかねっこ2月_temp.xls"
echo   run_generate_akanecco_shift.bat generate --target "【統一書式】あかねっこ2月_temp.xls"
echo.
call :run_with_log --help
call :maybe_pause
exit /b 0

:run_with_log
echo Log file: %LOG_PATH%
powershell -NoProfile -ExecutionPolicy Bypass -File "%RUNNER_PATH%" -ExePath "%EXE_PATH%" -LogPath "%LOG_PATH%" %*
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo Execution failed. See log file.
    echo Log file: %LOG_PATH%
)
exit /b %EXIT_CODE%

:maybe_pause
if "%NO_PAUSE%"=="1" exit /b 0
pause
exit /b 0