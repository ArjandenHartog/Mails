@echo off
python --version
IF ERRORLEVEL 1 (
  echo "Python is not geïnstalleerd. Installeer Python en probeer opnieuw."
  exit /b 1
)
pip install --user pandas pywin32
python lampentotaal_mail.py
