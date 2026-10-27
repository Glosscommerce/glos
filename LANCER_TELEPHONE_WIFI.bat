@echo off
setlocal
cd /d "%~dp0"
title MEMOIRE CORRECTEUR PRO gloss - acces telephone
if not exist ".venv\Scripts\python.exe" (
  echo Application tsy mbola voa-install. Alefaso aloha INSTALLER_WINDOWS.bat.
  pause
  exit /b 1
)
if not exist ".venv\dependencies_ready.txt" (
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 (pause & exit /b 1)
  echo ready>".venv\dependencies_ready.txt"
)
start "" cmd /c "timeout /t 5 /nobreak >nul & start http://localhost:8502"
echo.
echo Ny telefaonina dia tsy maintsy mifandray amin'ilay Wi-Fi mitovy amin'ny PC.
echo Sokafy ipconfig amin'ny CMD hafa raha hitady ny IPv4 Address an'ny PC.
echo Avy eo soraty amin'ny Chrome amin'ny telefaonina: http://ADRESSE-IP:8502
echo Aza akatona ity varavarankely ity raha mbola mampiasa ilay app.
echo.
".venv\Scripts\python.exe" -m streamlit run app\app.py --server.port 8502 --server.address 0.0.0.0 --server.headless true
pause
