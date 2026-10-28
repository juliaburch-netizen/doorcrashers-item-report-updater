@echo off
REM Drag the Item Report and the Doorcrashers file onto this file.
cd /d "%~dp0"
where py >nul 2>nul && (py "doorcrashers_updater.py" %*) || (python "doorcrashers_updater.py" %*)
echo.
pause
