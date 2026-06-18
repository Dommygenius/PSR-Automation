@echo off
cd /d "%~dp0"
set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"
if not exist "%PY%" set "PY=python"
echo.
echo PSR raw data export - uses Excel (2-5 min on large .xlsb). Please wait...
echo.
"%PY%" -u extract_raw_data.py --method com %*
if errorlevel 1 (
  echo.
  echo Excel method failed - trying pyxlsb fallback...
  "%PY%" -u extract_raw_data.py --method pyxlsb %*
)
