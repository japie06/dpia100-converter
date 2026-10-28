@echo off
REM Drag and drop any .csv file onto this .bat file to process it.
REM Keep this file in the same folder as dpia100.ps1.
REM -NoPause: this launcher has its own pause below, so the script must not ask twice.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dpia100.ps1" -InputCsvPath "%~1" -NoPause
echo.
pause
