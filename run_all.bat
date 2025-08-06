@echo off
setlocal

set "DEFAULT_MONTH=data\reports\2025-07"
set /p "MONTH=Monatsordner [%DEFAULT_MONTH%]: "
if "%MONTH%"=="" set "MONTH=%DEFAULT_MONTH%"

set "DEFAULT_LIST=Liste.xlsx"
set /p "LIST=Pfad zur Liste [%DEFAULT_LIST%]: "
if "%LIST%"=="" set "LIST=%DEFAULT_LIST%"

set "DEFAULT_OUT=report.csv"
set /p "OUTPUT=Ausgabedatei [%DEFAULT_OUT%]: "
if "%OUTPUT%"=="" set "OUTPUT=%DEFAULT_OUT%"

python -m dispatch.main run-all "%MONTH%" "%LIST%" --output "%OUTPUT%"

echo %DATE% %TIME% - run_all.bat ausgeführt mit "%MONTH%" "%LIST%" "%OUTPUT%" >> "logs\arbeitsprotokoll.md"

endlocal
