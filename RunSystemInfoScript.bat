@echo off

:: Check if running with elevated privileges (Admin rights)
NET SESSION >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    goto :run_ps1
) else (
    :: Re-run the script with Admin rights
    powershell -Command "Start-Process '%~0' -Verb RunAs"
    exit /b
)

:run_ps1
:: Set the working directory to the batch file's location
cd /d "%~dp0"

:: Execute the PowerShell script with bypassed execution policy
powershell -ExecutionPolicy Bypass -File "FinalHTMLTODB.ps1"

:: Pause the prompt to see any errors or output
pause
