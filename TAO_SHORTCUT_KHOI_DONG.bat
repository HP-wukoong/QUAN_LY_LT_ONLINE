@echo off
chcp 65001 >nul
rem Tao shortcut tren Desktop va (tuy chon) tu chay cung Windows
cd /d "%~dp0"
set "TARGET=%~dp0CHAY_TOOL.bat"
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Hoc Online O To.lnk');$s.TargetPath='%TARGET%';$s.WorkingDirectory='%~dp0';$s.WindowStyle=7;$s.Save()"
echo Da tao shortcut "Hoc Online O To" tren Desktop.
choice /M "Cho tool tu chay khi bat may (cua so thu nho)"
if errorlevel 2 goto end
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Startup')+'\Hoc Online O To.lnk');$s.TargetPath='%TARGET%';$s.WorkingDirectory='%~dp0';$s.WindowStyle=7;$s.Save()"
echo Da bat tu khoi dong cung Windows.
:end
pause
