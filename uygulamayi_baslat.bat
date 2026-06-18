@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Python sanal ortami bulunamadi.
  echo Once README.md icindeki backend kurulum adimlarini uygulayin.
  pause
  exit /b 1
)

start "AS9102 Teknik Resim Balonlama" cmd /k ""%~dp0.venv\Scripts\python.exe" -m streamlit run "%~dp0app.py""
timeout /t 5 /nobreak >nul
start "" "http://localhost:8501"
echo Streamlit uygulamasi baslatildi. Acilan terminal penceresini kapatmayin.
timeout /t 3 /nobreak >nul
