# Dispatch

Automatisiert die Auswertung von täglichen Anrufberichten der Servicetechniker und überträgt die Ergebnisse in `Liste.xlsx`.

## Voraussetzungen

- Python 3.11 oder neuer
- Pakete: pandas, openpyxl (`dispatch/requirements.txt`)

## Kurzanleitung

1. Lege die Tagesberichte in `data/reports/<YYYY-MM>/<TT>/` ab (z. B. `data/reports/2025-07/01`) mit Dateien `*7*.xlsx` und optional `*19*.xlsx`.
2. Führe `run_all.bat` aus und beantworte die abgefragten Pfade.
3. Die Ergebnisse werden in `Liste.xlsx` geschrieben und unter `logs/` protokolliert.

Optional kann der Tagesordner mit `python -m dispatch.create_day_dir` automatisch erstellt werden.

## Daten und Ergebnisse

- Die Ordnerstruktur `data/reports/` enthält Monatsordner (`YYYY-MM`) mit Tagesunterordnern (`TT`).
- Während der Verarbeitung entstehen Dateien wie `analysis.csv` oder `techniker_export.csv`, die nicht versioniert werden.

