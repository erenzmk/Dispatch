# Dispatch

Automatisiert die Auswertung von täglichen Anrufberichten der Servicetechniker und überträgt die Ergebnisse in `Liste.xlsx`.

## Voraussetzungen

- Python 3.11 oder neuer
- Abhängigkeiten laut `pyproject.toml` oder `requirements.txt`
- Basisinstallation: `pip install -r requirements.txt`
- Optionale GUI: `pip install --extra-index-url https://PySimpleGUI.net/install '.[gui]'`

## Kurzanleitung

1. Lege die Tagesberichte in `data/reports/<YYYY-MM>/<TT>/` ab (z. B. `data/reports/2025-07/01`) mit Dateien `*7*.xlsx` (Standard, über `--morning-pattern` konfigurierbar) und optional `*19*.xlsx`.
2. Starte die grafische Oberfläche mit `python run_all_gui.py`.
3. Wähle einen Tag oder den gesamten Monat aus. Die Ergebnisse werden in `Liste.xlsx` geschrieben und unter `logs/` protokolliert.

## Entwicklungsumgebung

- Virtuelle Umgebung: `python -m venv .venv` und Aktivierung mit `source .venv/bin/activate` (Linux/Mac) bzw. `\.venv\\Scripts\\activate` (Windows).
- Abhängigkeiten installieren: `pip install -r requirements.txt` oder für Entwicklung `pip install -e .`.
- Optionale GUI-Unterstützung: `pip install --extra-index-url https://PySimpleGUI.net/install '.[gui]'`.

## Testausführung

Die Tests laufen mit [pytest](https://pytest.org):

```bash
pytest
```

## Verzeichnisstruktur

- `dispatch/` – Quellcode und Tests (`dispatch/tests/`).
- `data/reports/` – Tagesberichte (Produktionsdaten, nicht versioniert).
- `results/` – Ausgabedateien (nicht versioniert).
- `logs/` – Protokolle (nicht versioniert).

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
- Bei Tests kommen ausschließlich synthetische Beispieldaten zum Einsatz.
- Produktionsdaten verbleiben lokal in `data/reports/` und werden aus Datenschutzgründen nicht versioniert.

