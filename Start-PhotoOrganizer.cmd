@echo off
set "INSTALLDIR=%~dp0"
set "APPDIR=%INSTALLDIR%App\"
set "LOCAL_LNK=%INSTALLDIR%Start-PhotoOrganizer.lnk"

if exist "%LOCAL_LNK%" (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "$install=$env:INSTALLDIR; $app=$env:APPDIR; $lnk=$env:LOCAL_LNK; try { $shell=New-Object -ComObject WScript.Shell; $shortcut=$shell.CreateShortcut($lnk); $shortcut.TargetPath=(Join-Path $install 'Start-PhotoOrganizer.cmd'); $shortcut.WorkingDirectory=$install; $icon=(Join-Path $app 'Assets\UDMRS-PhotoOrganizer.ico'); if (Test-Path -LiteralPath $icon) { $shortcut.IconLocation=($icon + ',0') }; $shortcut.Save() } catch { }" >nul 2>&1
)

where pwsh >nul 2>&1

if %errorlevel%==0 (
    pwsh -ExecutionPolicy Bypass -NoProfile -File "%APPDIR%PhotoOrganizerDashboard.ps1"
) else (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%APPDIR%PhotoOrganizerDashboard.ps1"
)
