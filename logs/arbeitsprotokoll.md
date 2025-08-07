## 2025-08-03
- Repository erkundet und keine AGENTS.md gefunden.
- Datei process_calls.py mit Verarbeitung der Techniker-Reports erstellt.
- Test tests/test_process_calls.py erstellt.
- requirements.txt um pandas erweitert.
- pytest ausgeführt: alle 23 Tests bestanden.
- Änderungen in Git committet.
- CLI-Test mit Beispielreport durchgeführt und temporäre Dateien entfernt.

## 2025-08-03 (später)
- Fehlerbehandlung für fehlende Berichtsdateien in process_calls.py ergänzt.
- tests/test_process_calls.py um fehlende Importe und Test für nicht vorhandene Datei erweitert.
- pytest ausgeführt.
- CLI mit Beispieldatei getestet und temporäre Dateien entfernt.
- CLI-Aufruf mit fehlender Datei getestet, erwartete Fehlermeldung erhalten.

## 2025-08-04
- Mergekonflikte geprüft und beseitigt.
- Abhängigkeiten installiert und Tests erneut ausgeführt.
- CLI mit Beispielreport getestet und temporäre Dateien entfernt.

## 2025-08-04 (später)
- Unbenutzten Import in den Tests entfernt.
- Abhängigkeiten erneut installiert und Tests erfolgreich ausgeführt.
- CLI-Test mit Beispieldatei durchgeführt und erzeugte Dateien gelöscht.

## 2025-08-05
- Repository geprüft und keine Mergekonflikte mehr gefunden.
- Abhängigkeiten installiert und Tests erfolgreich ausgeführt.
- CLI-Skript mit Beispieldatei getestet und temporäre Dateien entfernt.

## 2025-08-06
- Repository erneut geprüft und keine Konflikte festgestellt.
- Abhängigkeiten installiert und Tests erfolgreich ausgeführt.
- CLI mit Beispieldatei getestet und temporäre Dateien entfernt.
- CLI-Aufruf mit fehlender Datei getestet, erwartete Fehlermeldung erhalten.

## 2025-08-07
- Tests erneut erfolgreich ausgeführt.
- CLI mit generierter Beispieldatei getestet und Ergebnis gespeichert.
- CLI-Aufruf mit fehlender Datei liefert erwartete Fehlermeldung.
- Temporäre Dateien entfernt.

## 2025-08-08
- Repository erneut auf Konflikte geprüft, keine gefunden.
- Tests mit `pytest -q` ausgeführt: alle 24 Tests bestanden.
- CLI mit Beispieldatei und fehlender Datei getestet, temporäre Dateien anschließend entfernt.

## 2025-08-09
- Repository erneut geprüft, keine Konflikte.
- Tests mit `pytest -q` ausgeführt: alle 24 Tests bestanden.
- CLI mit Beispieldatei und fehlender Datei getestet, keine Ausgabedateien erstellt.

## 2025-08-10
- Unbekannte Techniker werden nun mit Vorschlägen protokolliert (`logs/unknown_technicians.log`).
- Tests erweitert und mit `pytest -q` ausgeführt.
- Monatsverarbeitung für `data/Juli_25` mit `data/Liste.xlsx` getestet; Logdatei erzeugt.
- `.gitignore` um die neue Logdatei ergänzt.
## 2025-08-10 (später, PowerShell)
- `pytest -q` in PowerShell ausgeführt: 25 Tests bestanden, 1 übersprungen.
- Versuch, `process_month` mit Here-Doc auszuführen, scheiterte wegen PowerShell-Syntax.
- Hinweis: in PowerShell `python -c` oder Here-String verwenden.
- `process_month` lässt sich direkt mit relativen Pfaden im Repo aufrufen:
- Keine Nutzung von `C:/Temp` notwendig; alle Pfade bleiben im Repo.

## 2025-08-11
- Mergekonflikte in Quell- und Testdateien beseitigt.
- `gather_valid_names` und GUI-Argumente bereinigt.
- `pytest` ausgeführt: alle Tests bestanden.

## 2025-08-04 (noch später)
- Spaltenzuordnung in `update_liste` korrigiert.
- Tests angepasst und mit `pytest -q` ausgeführt.
- CLI mit Beispieldatei getestet und temporäre Dateien entfernt.

## 2025-08-12
- Mergekonflikt in `tests/test_process_calls.py` aufgelöst und Parametrisierung kombiniert.
- `pytest tests/test_process_calls.py` ausgeführt: 6 Tests bestanden.

## 2025-08-05 (noch später)
- Datei `dispatch/__pycache__/process_reports.cpython-313.pyc` aus dem Repository entfernt.
- Verzeichnis `dispatch/__pycache__/` aus der Versionskontrolle entfernt.
- `.gitignore` geprüft: `__pycache__/` und `*.pyc` bereits vorhanden.

## 2025-08-05 (spät)
- Batchdatei `run_all.bat` erstellt, die Pfade abfragt und den Lauf protokolliert.
- README um Hinweise zur Batchdatei ergänzt.

## 2025-08-05 (noch später)
- README auf wesentliche Abschnitte reduziert.
- Verweise auf GUI, Unterbefehle und Debug-Optionen entfernt.
- `pytest` ausgeführt.
- Änderungen in Git committet.

## 2025-08-05 (sehr spät)
- Modul `aggregate_warnings.py` gelöscht und Eintrag in `__init__.py` entfernt.
- Funktion `log_unknown_technician` samt Aufruf in `process_reports.py` entfernt.
- Tests zu Warnungen gelöscht und Dokumentation angepasst.
- `pytest -q` ausgeführt.

## 2025-08-06 (Batchdatei automatisiert)
- run_all.bat auf feste Pfade ohne Abfragen umgestellt.
- Testlauf ausgeführt; Fehler wegen fehlender Daten, aber keine Eingabeaufforderungen.

## 2025-08-06 (Techniker-ID-Mapping)
- Modul `technicians.py` mit Funktion `load_id_map` erstellt.
- Tests hinzugefügt und mit `pytest -q` ausgeführt.
- Keine zusätzlichen Dateien angefallen.
## 2025-08-06 (summarize_by_id)
- Skript `summarize_by_id.py` erstellt, das Berichte nach Techniker-IDs auswertet.
- Test `test_summarize_by_id.py` hinzugefügt.
- `pytest -q` ausgeführt: alle 33 Tests bestanden.
- CLI mit Beispieldateien getestet und temporäre Dateien entfernt.

## 2025-08-06 (Ergebnisdatei)
- Erfolgreich Datei `results/2025-08-06.csv` mit Spalten ID, Techniker, Neu, Alt, Summe erstellt.
- Temporäre Dateien nach `to_delete/` verschoben.

## 2025-08-06 (Tests für summarize_by_id)
- Kleine Excel-Testdateien für neue, alte und unbekannte IDs angelegt.
- Neue Testdatei `dispatch/tests/test_summarize_by_id.py` erstellt.
- `pytest` ausgeführt: alle 33 Tests bestanden.

## 2025-08-06 (run_all-Filter)
- `run_all.bat` durchsucht nun nur Dateien mit `7` im Namen (`*7*.xlsx`).
- Kommentar im Skript angepasst.
- `pytest` ausgeführt: alle Tests bestanden.

## 2025-08-06 (run_all-Logging)
- `run_all.bat` legt das Verzeichnis `logs` bei Bedarf an.
- Nach jeder `summarize-id`-Ausführung werden Report- und Ergebnisdateien protokolliert.
- `pytest -q` ausgeführt: 33 Tests bestanden.

## 2025-08-06 (README: summarize-id)
- README um Abschnitt zur Nutzung von `python -m dispatch.main summarize-id` ergänzt.
- Beispielaufruf mit Ausgabe nach `results/` dokumentiert.
- `pytest -q` ausgeführt: alle Tests bestanden.
- CLI `summarize-id` mit Beispieldatei ausgeführt und erzeugte Datei entfernt.

## 2025-08-06 (Ergebnisdatei entfernt)
- Datei `results/2025-08-06.csv` gelöscht.
- `.gitignore` um `results/*.csv` erweitert.
- `pytest -q` ausgeführt: alle Tests bestanden.
## 2025-08-06 (CLI summarize-id Test)
- Neue Testdatei `dispatch/tests/test_main_summarize_id.py` erstellt.
- CLI-Unterbefehl `summarize-id` mit Beispieldateien getestet und CSV-Inhalte verifiziert.
- `pytest -q` ausgeführt: 34 Tests bestanden.
- Temporäre Dateien entfernt.
## 2025-08-06 (results-Verzeichnis und Ausgabeordner)
- `run_all.bat` legt das Verzeichnis `results` an, falls es fehlt.
- `dispatch/summarize_by_id.py` erzeugt den Ausgabeordner vor dem Schreiben der CSV.
- `pytest -q` ausgeführt: alle 34 Tests bestanden.
## 2025-08-06 (GUI als Einstiegsskript)
- `run_all.bat` nach `archive/` verschoben.
- README auf Start mit `python run_all_gui.py` angepasst.

## 2025-08-06 (PySimpleGUI)
- `PySimpleGUI` zu `dispatch/requirements.txt` hinzugefügt.
- README-Voraussetzungen um Paket und Installationshinweis ergänzt.
- `pip install -r dispatch/requirements.txt` ausgeführt.
- `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (PySimpleGUI 4.x Fix)
- PySimpleGUI-Version in `dispatch/requirements.txt` auf `~=4.60` beschränkt.
- README um Hinweis auf PySimpleGUI 4.x und Installationsbefehl mit privatem Index erweitert.
- Hinweis zur fehlenden 5.x-Unterstützung in `run_all_gui.py` ergänzt.
- `pip install --extra-index-url https://PySimpleGUI.net/install -r dispatch/requirements.txt` ausgeführt: Paket 4.x nicht gefunden.
- `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (PySimpleGUI 5.x)
- PySimpleGUI-Anforderung auf Version 5.x umgestellt.
- README-Voraussetzungen und Installationsbefehl aktualisiert.
- Hinweis in `run_all_gui.py` an PySimpleGUI 5.x angepasst.
- `pip install -r dispatch/requirements.txt` ausgeführt.
- `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (privater PySimpleGUI-Index)
- `dispatch/requirements.txt` mit `--extra-index-url` erweitert und `PySimpleGUI`-Zeile direkt darunter verschoben.
- `pip install -r dispatch/requirements.txt` ausgeführt.
- `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (README: privater Index)
- README-Voraussetzungen um Hinweis auf privaten PySimpleGUI-Index ergänzt.
- Installationsbefehl auf `pip install --extra-index-url https://PySimpleGUI.net/install -r dispatch/requirements.txt` geändert.
- `pip install --extra-index-url https://PySimpleGUI.net/install -r dispatch/requirements.txt` ausgeführt.
- `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (PySimpleGUI Theme-Check)
- In `run_all_gui.py` wird nach dem Import von PySimpleGUI geprüft, ob `sg` eine Funktion `theme` besitzt.
- Falls nicht, wird eine verständliche Fehlermeldung mit Hinweis auf den privaten Index ausgegeben und das Programm beendet.
- `python -m py_compile run_all_gui.py` und `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (Platzhalterpaket)
- requirements.txt, README und run_all_gui.py angepasst.
- Fehler wurde durch ein Platzhalterpaket ausgelöst; künftig wird der private Index genutzt.

## 2025-08-06 (safe_load_workbook in summarize_by_id)
- Direkten Import von `load_workbook` entfernt und `safe_load_workbook` aus `process_reports` eingebunden.
- Aufruf in `summarize_by_id.py` angepasst.
- `pip install -r dispatch/requirements.txt` ausgeführt.
- `pytest -q` ausgeführt: 34 Tests bestanden.
## 2025-08-06 (process vor summarize-id)
- `summarize_day` ruft jetzt vor der Zusammenfassung `dispatch.main process` mit `day_dir` und `liste` auf.
- Arbeitsprotokoll erweitert und Code getestet.
- `pytest -q` ausgeführt: alle Tests bestanden.

## 2025-08-06 (run_gui Ausgabepfad)
- GUI um Eingabefeld für den Ausgabepfad ergänzt.
- `process_month` verwendet den gewählten Pfad oder einen Standardnamen mit Monat.
- `python -m py_compile run_all_gui.py` und `pytest -q` ausgeführt: 34 Tests bestanden.

## 2025-08-06 (Monatsblatt-Prüfung)
- `main` und `update_liste` legen fehlende Monatsblätter automatisch mit Kopfzeile an.
- Tests für das neue Verhalten ergänzt.
- `pytest -q` ausgeführt: 36 Tests bestanden.

## 2025-08-06 (GUI-Fehlerbehandlung)
- subprocess-Aufrufe in `summarize_day` und `process_month` mit `try`/`except` abgesichert.
- Fehler werden ins Arbeitsprotokoll geschrieben und als Popup angezeigt.
- GUI zeigt nur bei Erfolg eine Abschlussmeldung.
- `python -m py_compile run_all_gui.py` und `pytest -q` ausgeführt: 36 Tests bestanden.
- Keine zusätzlichen Dateien entstanden.

## 2025-08-06 (Techniker-ID-Mapping überspringt Titelzeile)
- `load_id_map` sucht nun nach der Kopfzeile statt nur Zeile 1 zu verwenden.
- Zusätzlichen Test für Tabellen mit Titelzeile hinzugefügt.
- `pytest -q` ausgeführt: alle Tests bestanden.

## 2025-08-06 (später)
- `load_calls` verwirft nun unbekannte Techniker statt sie zu zählen.
- `update_liste` fügt neue Namen nicht mehr hinzu, sondern protokolliert sie.
- Tests angepasst und mit `pytest -q` ausgeführt: 39 Tests bestanden.

## 2025-08-06 (Filter und Duplikate)
- `load_calls` verarbeitet nur noch Blätter mit `Report` im Namen.
- Mehrfache `Work Order Number` werden ignoriert, sodass jeder Auftrag nur einmal zählt.
- Zwei neue Tests prüfen Filter- und Duplikaterkennung.
- `pytest -q` ausgeführt: 41 Tests bestanden.

## 2025-08-06 (Blattnamen)
- `RELEVANT_SHEET_PATTERNS` um "West Central" und "Detailed" erweitert.
- `load_calls` verarbeitet bei leerer Filterliste alle Blätter.
- Testfall für ein Blatt "West Central" ergänzt.
- `pytest -q` ausgeführt: 43 Tests bestanden.

## 2025-08-07 (relevante Blätter)
- `load_calls` meldet nun fehlende passende Arbeitsblätter und nennt gesuchte Muster sowie vorhandene Blattnamen.
- Test `test_load_calls_reports_missing_relevant_sheets` hinzugefügt.
- `pytest -q` ausgeführt: 44 Tests bestanden.

## 2025-08-07 (Reportdaten entfernt)
- Alle Excel-Dateien unter `data/reports` aus dem Repository gelöscht.
- `.gitignore` erweitert, damit keine echten Reports mehr eingecheckt werden.
- `pytest -q` ausgeführt: 44 Tests bestanden.

## 2025-08-07 (Morgenreport-Muster)
- Morgenreport-Muster über `DEFAULT_MORNING_PATTERN` und Option `--morning-pattern` konfigurierbar.
- Fallback nutzt erste `.xlsx`-Datei und listet vorhandene Dateien auf.
- Tests für alternative Dateinamen und Fallback ergänzt.
- `pytest -q` ausgeführt: 46 Tests bestanden.

## 2025-08-07 (Subprozess-Fehlerausgabe)
- subprocess.run-Aufrufe in run_all_gui.py erfassen nun Ausgabe und Fehler.
- Bei Fehlermeldungen werden stdout und stderr ins Protokoll geschrieben.
- Test `test_process_month_logs_subprocess_failure` hinzugefügt.
- `pytest -q` ausgeführt: 47 Tests bestanden.


## 2025-08-07 (Call-Listen)

- Funktion `extract_calls_by_id` in `process_reports.py` implementiert.
- `summarize_by_id.summarize_report` ruft nun diese Funktion auf und gibt die Call-Listen zurück.
- `run_all_gui.summarize_day` und `process_month` protokollieren die Call-Listen optional.
- Test `test_summarize_by_id.py` um Call-Listen erweitert und neuer Test für `extract_calls_by_id` hinzugefügt.

## 2025-08-07 (strukturierte Protokolle)
- `_log` akzeptiert optionale strukturierte Daten und hängt sie formatiert an.
- `summarize_day` und `process_month` protokollieren Call-Listen pro Techniker-ID bei Erfolg.
- Test für erfolgreiche Monatsverarbeitung ergänzt, der Call-Listen im Protokoll erwartet.
- Überflüssige `.log`-Dateien entfernt und `.gitignore` erweitert.

2025-08-07 15:14:56 - dispatch.main process mit "data\reports\2025-07\01" "Liste.xlsx"
2025-08-07 15:14:58 - Report "data\reports\2025-07\01\19 Uhr.xlsx" -> "results\01_19 Uhr_summary.csv"
2025-08-07 15:15:00 - Report "data\reports\2025-07\01\7 Uhr.xlsx" -> "results\01_7 Uhr_summary.csv"
2025-08-07 15:15:00 - Call-Listen für Tag "01"
{}
2025-08-07 15:15:02 - dispatch.main process mit "data\reports\2025-07\02" "Liste.xlsx"
2025-08-07 15:15:04 - Report "data\reports\2025-07\02\02.07.25 0705.xlsx" -> "results\02_02.07.25 0705_summary.csv"
2025-08-07 15:15:05 - Report "data\reports\2025-07\02\02.07.25 1900.xlsx" -> "results\02_02.07.25 1900_summary.csv"
2025-08-07 15:15:05 - Call-Listen für Tag "02"
{}
2025-08-07 15:15:08 - dispatch.main process mit "data\reports\2025-07\03" "Liste.xlsx"
2025-08-07 15:15:10 - Report "data\reports\2025-07\03\03.07.25 0720.xlsx" -> "results\03_03.07.25 0720_summary.csv"
2025-08-07 15:15:11 - Report "data\reports\2025-07\03\03.07.25 1900.xlsx" -> "results\03_03.07.25 1900_summary.csv"
2025-08-07 15:15:11 - Call-Listen für Tag "03"
{}
2025-08-07 15:15:13 - dispatch.main process mit "data\reports\2025-07\04" "Liste.xlsx"
2025-08-07 15:15:15 - Report "data\reports\2025-07\04\04.07.25 0718.xlsx" -> "results\04_04.07.25 0718_summary.csv"
2025-08-07 15:15:16 - Report "data\reports\2025-07\04\04.07.25 1901.xlsx" -> "results\04_04.07.25 1901_summary.csv"
2025-08-07 15:15:16 - Call-Listen für Tag "04"
{}
2025-08-07 15:15:18 - dispatch.main process mit "data\reports\2025-07\07" "Liste.xlsx"
2025-08-07 15:15:20 - Report "data\reports\2025-07\07\07.07.25 0727.xlsx" -> "results\07_07.07.25 0727_summary.csv"
2025-08-07 15:15:21 - Report "data\reports\2025-07\07\07.07.25 1900.xlsx" -> "results\07_07.07.25 1900_summary.csv"
2025-08-07 15:15:21 - Call-Listen für Tag "07"
{}
2025-08-07 15:15:24 - dispatch.main process mit "data\reports\2025-07\08" "Liste.xlsx"
2025-08-07 15:15:25 - Report "data\reports\2025-07\08\08.07.25 0701.xlsx" -> "results\08_08.07.25 0701_summary.csv"
2025-08-07 15:15:27 - Report "data\reports\2025-07\08\08.07.25 1901.xlsx" -> "results\08_08.07.25 1901_summary.csv"
2025-08-07 15:15:27 - Call-Listen für Tag "08"
{}
2025-08-07 15:15:28 - dispatch.main process mit "data\reports\2025-07\09" "Liste.xlsx"
2025-08-07 15:15:30 - Report "data\reports\2025-07\09\09.07.25 0718.xlsx" -> "results\09_09.07.25 0718_summary.csv"
2025-08-07 15:15:32 - Report "data\reports\2025-07\09\09.07.25 1901.xlsx" -> "results\09_09.07.25 1901_summary.csv"
2025-08-07 15:15:32 - Call-Listen für Tag "09"
{}
2025-08-07 15:15:34 - dispatch.main process mit "data\reports\2025-07\10" "Liste.xlsx"
2025-08-07 15:15:36 - Report "data\reports\2025-07\10\10.07.25 0709.xlsx" -> "results\10_10.07.25 0709_summary.csv"
2025-08-07 15:15:37 - Report "data\reports\2025-07\10\10.07.25 1901.xlsx" -> "results\10_10.07.25 1901_summary.csv"
2025-08-07 15:15:37 - Call-Listen für Tag "10"
{}
2025-08-07 15:15:39 - dispatch.main process mit "data\reports\2025-07\11" "Liste.xlsx"
2025-08-07 15:15:41 - Report "data\reports\2025-07\11\11.07.25 0710.xlsx" -> "results\11_11.07.25 0710_summary.csv"
2025-08-07 15:15:42 - Report "data\reports\2025-07\11\11.07.25 1900.xlsx" -> "results\11_11.07.25 1900_summary.csv"
2025-08-07 15:15:42 - Call-Listen für Tag "11"
{}
2025-08-07 15:15:44 - dispatch.main process mit "data\reports\2025-07\14" "Liste.xlsx"
2025-08-07 15:15:46 - Report "data\reports\2025-07\14\14.07.25 0717.xlsx" -> "results\14_14.07.25 0717_summary.csv"
2025-08-07 15:15:47 - Report "data\reports\2025-07\14\14.07.25 1900.xlsx" -> "results\14_14.07.25 1900_summary.csv"
2025-08-07 15:15:47 - Call-Listen für Tag "14"
{}
2025-08-07 15:15:49 - dispatch.main process mit "data\reports\2025-07\15" "Liste.xlsx"
2025-08-07 15:15:51 - Report "data\reports\2025-07\15\15.07.25 0700.xlsx" -> "results\15_15.07.25 0700_summary.csv"
2025-08-07 15:15:53 - Report "data\reports\2025-07\15\15.07.25 1900.xlsx" -> "results\15_15.07.25 1900_summary.csv"
2025-08-07 15:15:53 - Call-Listen für Tag "15"
{}
2025-08-07 15:15:55 - dispatch.main process mit "data\reports\2025-07\16" "Liste.xlsx"
2025-08-07 15:15:57 - Report "data\reports\2025-07\16\16.07.25 1001.xlsx" -> "results\16_16.07.25 1001_summary.csv"
2025-08-07 15:15:58 - Report "data\reports\2025-07\16\16.07.25 1900.xlsx" -> "results\16_16.07.25 1900_summary.csv"
2025-08-07 15:15:58 - Call-Listen für Tag "16"
{}
2025-08-07 15:16:01 - dispatch.main process mit "data\reports\2025-07\17" "Liste.xlsx"
2025-08-07 15:16:03 - Report "data\reports\2025-07\17\17.07.25 0706.xlsx" -> "results\17_17.07.25 0706_summary.csv"
2025-08-07 15:16:06 - Report "data\reports\2025-07\17\17.07.25 1901.xlsx" -> "results\17_17.07.25 1901_summary.csv"
2025-08-07 15:16:06 - Call-Listen für Tag "17"
{}
2025-08-07 15:16:08 - dispatch.main process mit "data\reports\2025-07\18" "Liste.xlsx"
2025-08-07 15:16:10 - Report "data\reports\2025-07\18\18.07.25 0716.xlsx" -> "results\18_18.07.25 0716_summary.csv"
2025-08-07 15:16:12 - Report "data\reports\2025-07\18\18.07.25 1900.xlsx" -> "results\18_18.07.25 1900_summary.csv"
2025-08-07 15:16:12 - Call-Listen für Tag "18"
{}
2025-08-07 15:16:14 - dispatch.main process mit "data\reports\2025-07\21" "Liste.xlsx"
2025-08-07 15:16:16 - Report "data\reports\2025-07\21\21.07.25 0730.xlsx" -> "results\21_21.07.25 0730_summary.csv"
2025-08-07 15:16:18 - Report "data\reports\2025-07\21\21.07.25 1901.xlsx" -> "results\21_21.07.25 1901_summary.csv"
2025-08-07 15:16:19 - Call-Listen für Tag "21"
{}
2025-08-07 15:16:21 - dispatch.main process mit "data\reports\2025-07\22" "Liste.xlsx"
2025-08-07 15:16:23 - Report "data\reports\2025-07\22\22.07.25 0701.xlsx" -> "results\22_22.07.25 0701_summary.csv"
2025-08-07 15:16:24 - Report "data\reports\2025-07\22\22.07.25 1901.xlsx" -> "results\22_22.07.25 1901_summary.csv"
2025-08-07 15:16:24 - Call-Listen für Tag "22"
{}
2025-08-07 15:16:26 - dispatch.main process mit "data\reports\2025-07\23" "Liste.xlsx"
2025-08-07 15:16:28 - Report "data\reports\2025-07\23\23.07.25 0724.xlsx" -> "results\23_23.07.25 0724_summary.csv"
2025-08-07 15:16:31 - Report "data\reports\2025-07\23\23.07.25 1901.xlsx" -> "results\23_23.07.25 1901_summary.csv"
2025-08-07 15:16:31 - Call-Listen für Tag "23"
{}
2025-08-07 15:16:33 - dispatch.main process mit "data\reports\2025-07\24" "Liste.xlsx"
2025-08-07 15:16:35 - Report "data\reports\2025-07\24\24.07.25 0717.xlsx" -> "results\24_24.07.25 0717_summary.csv"
2025-08-07 15:16:37 - Report "data\reports\2025-07\24\24.07.25 1900.xlsx" -> "results\24_24.07.25 1900_summary.csv"
2025-08-07 15:16:37 - Call-Listen für Tag "24"
{}
2025-08-07 15:16:40 - Fehler bei dispatch.main process mit "data\reports\2025-07\25" "Liste.xlsx": Command '['python', '-m', 'dispatch.main', 'process', 'data\\reports\\2025-07\\25', 'Liste.xlsx']' returned non-zero exit status 1.
STDOUT: 
STDERR: Traceback (most recent call last):
  File "<frozen runpy>", line 198, in _run_module_as_main
  File "<frozen runpy>", line 88, in _run_code
  File "C:\Users\egencer\Documents\GitHub\Dispatch\dispatch\main.py", line 71, in <module>
    main()
    ~~~~^^
  File "C:\Users\egencer\Documents\GitHub\Dispatch\dispatch\main.py", line 53, in main
    process_reports.main([str(args.day_dir), str(args.liste)])
    ~~~~~~~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\egencer\Documents\GitHub\Dispatch\dispatch\process_reports.py", line 534, in main
    update_liste(args.liste, month_sheet, target_date, morning_summary)
    ~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\egencer\Documents\GitHub\Dispatch\dispatch\process_reports.py", line 373, in update_liste
    wb.save(liste)
    ~~~~~~~^^^^^^^
  File "C:\Users\egencer\AppData\Local\Programs\Python\Python313\Lib\site-packages\openpyxl\workbook\workbook.py", line 386, in save
    save_workbook(self, filename)
    ~~~~~~~~~~~~~^^^^^^^^^^^^^^^^
  File "C:\Users\egencer\AppData\Local\Programs\Python\Python313\Lib\site-packages\openpyxl\writer\excel.py", line 291, in save_workbook
    archive = ZipFile(filename, 'w', ZIP_DEFLATED, allowZip64=True)
  File "C:\Users\egencer\AppData\Local\Programs\Python\Python313\Lib\zipfile\__init__.py", line 1367, in __init__
    self.fp = io.open(file, filemode)
              ~~~~~~~^^^^^^^^^^^^^^^^
PermissionError: [Errno 13] Permission denied: 'Liste.xlsx'

2025-08-07 15:17:41 - dispatch.main process mit "data\reports\2025-07\28" "Liste.xlsx"
2025-08-07 15:17:42 - Report "data\reports\2025-07\28\28.07.25 0726.xlsx" -> "results\28_28.07.25 0726_summary.csv"
