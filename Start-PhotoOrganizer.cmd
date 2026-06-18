@echo off
set "INSTALLDIR=%~dp0"
set "APPDIR=%INSTALLDIR%App\"

where pwsh >nul 2>&1

if %errorlevel%==0 (
    pwsh -ExecutionPolicy Bypass -NoProfile -File "%APPDIR%PhotoOrganizerDashboard.ps1"
) else (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%APPDIR%PhotoOrganizerDashboard.ps1"
)
