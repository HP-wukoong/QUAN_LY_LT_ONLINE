@echo off
chcp 65001 >nul
title QUAN LY HOC ONLINE O TO - KHONG TAT CUA SO NAY (thu nho duoc)
cd /d "%~dp0"

rem --- Tim Python that (bo qua ban gia cua Microsoft Store) ---
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY (
  python -c "import sys" >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo.
  echo [LOI] May chua cai Python 3.
  echo  - Tai Python 3.11 tai https://www.python.org/downloads/
  echo  - Khi cai nho tick "Add python.exe to PATH"
  echo.
  pause
  exit /b
)

rem openpyxl da kem san trong app\vendor - khong can mang/pip
%PY% app\server.py
if errorlevel 1 (
  echo.
  echo [LOI] Tool dung bat thuong. Chup man hinh cua so nay gui nguoi ho tro.
  pause
)
