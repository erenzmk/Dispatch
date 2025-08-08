@echo off
rem MONTH: Pfad zum Monatsordner mit Tagesreports (z. B. data\reports\2025-07)
set "MONTH=data\reports\2025-07"
rem LIST: Pfad zur Technikerliste
set "LIST=Liste.xlsx"
rem OUTPUT: Datei für die Monatsanalyse
set "OUTPUT=report.csv"

rem sicherstellen, dass das Logs-Verzeichnis existiert
if not exist "logs" mkdir "logs"

rem Monatsdaten verarbeiten und analysieren
python -m dispatch.main run-all "%MONTH%" "%LIST%" --output "%OUTPUT%"

rem sicherstellen, dass das Results-Verzeichnis existiert
if not exist "results" mkdir "results"

rem Jeden Tagesreport mit "7" im Dateinamen nach Techniker-ID zusammenfassen
for /R "%MONTH%" %%F in (*7*.xlsx) do (
    for %%D in ("%%~dpF.") do (
        rem %%~nD: Tagesordner, %%~nF: Reportname
        python -m dispatch.main summarize-id "%%F" "%LIST%" --output "results\%%~nD_%%~nF_summary.csv"
        echo %DATE% %TIME% - Report "%%F" -> "results\%%~nD_%%~nF_summary.csv" >> "logs\arbeitsprotokoll.md"
    )
)

rem Ausführung protokollieren
echo %DATE% %TIME% - run_all.bat ausgeführt mit "%MONTH%" "%LIST%" "%OUTPUT%" >> "logs\arbeitsprotokoll.md"
