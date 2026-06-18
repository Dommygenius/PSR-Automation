@echo off
cd /d "%~dp0"
echo.
echo PSR raw data export - uses Excel (2-5 min on large .xlsb). Please wait...
echo.
python extract_raw_data.py --method com %*
if errorlevel 1 (
  echo.
  echo Excel method failed - trying pyxlsb fallback...
  python extract_raw_data.py --method pyxlsb %*
)
