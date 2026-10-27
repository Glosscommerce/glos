@echo off
setlocal
cd /d "%~dp0"
title MEMOIRE CORRECTEUR PRO gloss

if not exist ".venv\pyvenv.cfg" (
  if exist ".venv" rmdir /s /q ".venv"
  where py >nul 2>nul
  if errorlevel 1 (
    echo Python tsy hita. Apetraho aloha ny Python 3.11 na vaovao kokoa.
    echo https://www.python.org/downloads/windows/
    pause
    exit /b 1
  )
  py -3 -m venv .venv
  if errorlevel 1 (echo Tsy nahomby ny famoronana .venv.& pause & exit /b 1)
)
if not exist ".venv\Scripts\python.exe" (
  echo Tsy feno ny .venv. Fafao ilay dossier .venv ary avereno.
  pause
  exit /b 1
)
if not exist ".venv\dependencies_ready.txt" (
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  if errorlevel 1 (echo Tsy afaka nampiditra pip.& pause & exit /b 1)
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 (echo Tsy nahomby ny fametrahana dependances.& pause & exit /b 1)
  echo ready>".venv\dependencies_ready.txt"
)
".venv\Scripts\python.exe" -m streamlit run app\app.py --server.port 8502 --server.address 127.0.0.1 --server.headless false
pause
