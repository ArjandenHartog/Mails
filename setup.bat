@echo off
echo LampenTotaal Mail Setup
echo =====================

REM Check for Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python niet gevonden. Open Microsoft Store om Python te installeren...
    start ms-windows-store://pdp/?ProductId=9PJPW5LDXLZB
    echo Herstart deze setup nadat Python is geinstalleerd.
    pause
    exit /b 1
)

REM Install/upgrade required packages
echo.
echo Installeren benodigde packages...
python -m ensurepip --user
python -m pip install --user --upgrade pip --quiet
python -m pip install --user --quiet pandas pywin32

REM Start application
echo.
echo Setup voltooid. LampenTotaal Mail wordt nu gestart...
start pythonw lampentotaal_mail.py

exit /b 0
