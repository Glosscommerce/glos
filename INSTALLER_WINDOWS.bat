@echo off
setlocal
cd /d "%~dp0"
title Installer - MEMOIRE CORRECTEUR PRO gloss
where powershell.exe >nul 2>nul
if errorlevel 1 (
  echo Windows PowerShell tsy hita. Tsy afaka manohy.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer.ps1"
if errorlevel 1 (
  echo.
  echo Nisy olana nandritra ny installation. Jereo ireo hafatra etsy ambony.
  pause
  exit /b 1
)
echo.
echo Installation vita. Tokony hisy raccourci eo amin'ny Bureau.
pause
