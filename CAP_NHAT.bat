@echo off
chcp 65001 >nul
title CAP NHAT TOOL HOC ONLINE O TO TU GITHUB
cd /d "%~dp0"
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY ( python -c "import sys" >nul 2>nul && set "PY=python" )
if not defined PY ( echo [LOI] May chua cai Python 3. & pause & exit /b )
echo Hay TAT cua so CHAY_TOOL truoc khi cap nhat.
pause
%PY% app\server.py --update
echo.
echo Xong. Mo lai CHAY_TOOL.bat de dung ban moi.
pause
