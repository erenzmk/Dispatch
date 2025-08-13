# Arbeitslog

## 2024-05-16
- `run_all_gui.py` in `run_dispatch.py` umbenannt und GUI-Code entfernt.
- Nur `summarize_day` und `process_month` mit `_log` beibehalten.
- Importpfade auf `run_dispatch` angepasst.
- Tests und `pyproject.toml` entsprechend aktualisiert.
## 2025-08-13
- `run_current_month.pyw` zu `run_current_month.py` umbenannt und Wrapper `run_current_month.pyw` für `start_dispatch.bat` erstellt.
- Import in `run_current_month.py` auf `from run_dispatch import process_month` sichergestellt.
- Optionale GUI-Abhängigkeiten und `dispatch-gui`-Entry-Point entfernt, README entsprechend angepasst und `pip install -e .` ausgeführt.
- CLI in `run_dispatch.py` eingeführt und Tests auf `run_dispatch`-Importe umgestellt.
- GUI-spezifische Mock-Aufrufe bereinigt und Testlauf mit `pytest` bestätigt.
## 2025-08-13 (2)
- Neues Skript `dispatch/write_liste.py` zum Eintragen der Tagesdaten in `Liste.xlsx` erstellt.
- Alias-Mappings aus "Technikernamen + PUDO" werden eingelesen, Tagesblöcke erkannt oder bei Bedarf ergänzt.
- Aggregationslogik für numerische und textuelle Felder umgesetzt.
## 2025-08-13 (3)
- `write_liste.py` vollständig auf Wochenblöcke umgestellt und Blocksuche über Datenzeilen implementiert.
- Schreiblogik nutzt kanonische Namen und Datum, ohne neue Blöcke oder Zeilen anzulegen.
