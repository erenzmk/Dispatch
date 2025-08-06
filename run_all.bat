@echo off
set "MONTH=data\reports\2025-07"
set "LIST=Liste.xlsx"
set "OUTPUT=report.csv"
python -m dispatch.main run-all "%MONTH%" "%LIST%" --output "%OUTPUT%"
echo %DATE% %TIME% - run_all.bat ausgeführt mit "%MONTH%" "%LIST%" "%OUTPUT%" >> "logs\arbeitsprotokoll.md"
