@echo off
cd /d "%~dp0"
set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"
if not exist "%PY%" set "PY=python"
if "%1"=="--into-xlsb" (
  echo Close the .xlsb in Excel/Cursor first. This may take 5-15 minutes.
)
"%PY%" -u extract_valuable_to_sheet.py %*
pause
