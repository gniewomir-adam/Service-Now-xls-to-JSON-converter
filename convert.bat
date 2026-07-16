@echo off

echo ======================================
echo Whisperer Excel to JSON Converter
echo ======================================
echo.

python whisperer_excel_to_json.py ^
  --input "DREAM - Monitoring status & Issues & Incidents Tracker.xlsx" ^
  --output whisperer_input.json ^
  --config whisperer_mapping_config.json ^
  --sheet "Issues Tracker" ^
  --log-level INFO

echo.
echo Exit code: %ERRORLEVEL%
echo.

pause