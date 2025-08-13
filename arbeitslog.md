# Arbeitslog

## 2024-05-16
- `run_all_gui.py` in `run_dispatch.py` umbenannt und GUI-Code entfernt.
- Nur `summarize_day` und `process_month` mit `_log` beibehalten.
- Importpfade auf `run_dispatch` angepasst.
- Tests und `pyproject.toml` entsprechend aktualisiert.
## 2025-08-13
- `run_current_month.pyw` zu `run_current_month.py` umbenannt und Wrapper `run_current_month.pyw` für `start_dispatch.bat` erstellt.
- Import in `run_current_month.py` auf `from run_dispatch import process_month` sichergestellt.
