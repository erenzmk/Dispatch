# Dispatch

Automatisiert die Auswertung von täglichen Anrufberichten der Servicetechniker und überträgt die Ergebnisse in `Liste.xlsx`.

## Voraussetzungen

- Python 3.11 oder neuer
- Pakete: pandas, openpyxl, PySimpleGUI (`dispatch/requirements.txt`)
- Installation: `pip install -r dispatch/requirements.txt`

## Kurzanleitung

1. Lege die Tagesberichte in `data/reports/<YYYY-MM>/<TT>/` ab (z. B. `data/reports/2025-07/01`) mit Dateien `*7*.xlsx` und optional `*19*.xlsx`.
2. Starte die grafische Oberfläche mit `python run_all_gui.py`.
3. Wähle einen Tag oder den gesamten Monat aus. Die Ergebnisse werden in `Liste.xlsx` geschrieben und unter `logs/` protokolliert.

## Auswertung nach Techniker-ID

Ein einzelner Tagesbericht lässt sich mit `summarize-id` nach Techniker-IDs zusammenfassen:

```bash
python -m dispatch.main summarize-id data/report.xlsx data/Liste.xlsx --output results/2025-08-06.csv
```

Die Ausgabedatei landet im Verzeichnis `results/`.

Optional kann der Tagesordner mit `python -m dispatch.create_day_dir` automatisch erstellt werden.

## Daten und Ergebnisse

- Die Ordnerstruktur `data/reports/` enthält Monatsordner (`YYYY-MM`) mit Tagesunterordnern (`TT`).
- Während der Verarbeitung entstehen Dateien wie `analysis.csv` oder `techniker_export.csv`, die nicht versioniert werden.

