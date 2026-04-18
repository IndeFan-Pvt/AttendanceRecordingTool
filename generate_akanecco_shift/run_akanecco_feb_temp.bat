@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%generate_akanecco_shift.exe"
set "RUNNER_PATH=%SCRIPT_DIR%run_with_utf8_log.ps1"
set "TARGET_PATH=%SCRIPT_DIR%..\..\【統一書式】あかねっこ2月_temp.xls"
set "LOG_DIR=%SCRIPT_DIR%logs"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TIMESTAMP=%%I"
set "LOG_PATH=%LOG_DIR%\run_akanecco_feb_temp_%TIMESTAMP%.log"

if not exist "%EXE_PATH%" (
    echo generate_akanecco_shift.exe was not found.
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[System.IO.File]::WriteAllLines('%LOG_PATH%', @('[%date% %time%] ERROR: generate_akanecco_shift.exe was not found.', '[%date% %time%] SCRIPT_DIR=%SCRIPT_DIR%'), [System.Text.UTF8Encoding]::new($false))"
    echo Log file: %LOG_PATH%
    call :maybe_pause
    exit /b 1
)

if not exist "%RUNNER_PATH%" (
    echo run_with_utf8_log.ps1 was not found.
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[System.IO.File]::WriteAllLines('%LOG_PATH%', @('[%date% %time%] ERROR: run_with_utf8_log.ps1 was not found.', '[%date% %time%] SCRIPT_DIR=%SCRIPT_DIR%'), [System.Text.UTF8Encoding]::new($false))"
    echo Log file: %LOG_PATH%
    call :maybe_pause
    exit /b 1
)

if not exist "%TARGET_PATH%" (
    echo Target workbook was not found.
    echo Expected path: %TARGET_PATH%
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[System.IO.File]::WriteAllLines('%LOG_PATH%', @('[%date% %time%] ERROR: target workbook was not found.', '[%date% %time%] TARGET_PATH=%TARGET_PATH%'), [System.Text.UTF8Encoding]::new($false))"
    echo Log file: %LOG_PATH%
    call :maybe_pause
    exit /b 1
)

echo Target file: %TARGET_PATH%
echo Log file: %LOG_PATH%
powershell -NoProfile -ExecutionPolicy Bypass -File "%RUNNER_PATH%" -ExePath "%EXE_PATH%" -LogPath "%LOG_PATH%" generate --target "%TARGET_PATH%"
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo Execution failed. See log file.
    echo Log file: %LOG_PATH%
)
call :maybe_pause
exit /b %EXIT_CODE%

:maybe_pause
if "%NO_PAUSE%"=="1" exit /b 0
pause
exit /b 0