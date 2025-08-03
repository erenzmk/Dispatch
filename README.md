# Dispatch

Dieses Repository automatisiert die Auswertung von täglichen Anrufberichten der Servicetechniker. Der Kern der Logik liegt im Paket `dispatch` und wird über die Befehle in `main.py` bereitgestellt. Die Skripte lesen die Excel-Reporte der Früh- und Spätschicht, fassen die Werte je Techniker zusammen und schreiben die Ergebnisse in `Liste.xlsx`.

## Nutzung

### Einzelnen Tag verarbeiten
```bash
python main.py process data/Juli_25/01.07 data/Liste.xlsx
```
Der Befehl erwartet den Ordner des Tages (z. B. `data/Juli_25/01.07`) und den Pfad zur Excel-Gesamtliste (`data/Liste.xlsx`).

### Kompletten Monat verarbeiten
```bash
python main.py process-month data/Juli_25 data/Liste.xlsx
```
Iteriert über alle Unterordner in `data/Juli_25` und verarbeitet jeden Tag, für den passende `*7*.xlsx`‑ und `*19*.xlsx`‑Dateien existieren.

### Monat auswerten
```bash
python main.py analyze data/Juli_25 data/Liste.xlsx --output report.csv
```
Erstellt eine CSV-Datei mit Kategorien wie `no_calls` und `region_mismatch`.

### Unbekannte Techniker zählen
```bash
python main.py warnings data/Juli_25 --liste data/Liste.xlsx
```
Durchsucht alle Berichte und fasst unbekannte Namen samt Häufigkeit zusammen.

### Alles in einem Schritt
```bash
python main.py run-all data/Juli_25 data/Liste.xlsx --output report.csv
```
Verarbeitet den Monat, erstellt die Analyse und zeigt unbekannte Techniker in einem einzelnen Durchlauf an.

## Generierte Dateien

Dateien wie `analysis.csv` oder `techniker_export.csv` werden zur Laufzeit erstellt und nicht versioniert.

## Fehlerbehebung

### Leere Auswertungen
Wenn keine Einträge für einen Techniker erscheinen, prüfe, ob der Tagesreport die Zeile mit `Employee ID` enthält. Fehlt sie, kann der Header nicht erkannt werden.

### Fehlende Dateien
Der Tagesordner muss eine Morgen-Datei (`*7*.xlsx`) und optional eine Abend-Datei (`*19*.xlsx`) enthalten. Andernfalls schlägt die Verarbeitung fehl.

### Falsche Blattnamen
In `Liste.xlsx` muss ein Arbeitsblatt nach dem Muster `<Monat>_<JJ>` existieren, z. B. `Juli_25`. Ein `KeyError` weist auf ein fehlendes oder falsch benanntes Blatt hin.

### Ausführliche Protokolle
Für detailliertere Ausgaben kann das Logging auf DEBUG gesetzt werden:
```bash
$env:LOGLEVEL="DEBUG"
python main.py process data/Juli_25/01.07 data/Liste.xlsx
```

### Zuordnung unbekannter Namen
Mit der kleinen Tk‑Oberfläche lassen sich unbekannte Namen interaktiv zuweisen:
```bash
python assign_gui.py  # nutzt standardmäßig data/ und Liste.xlsx
```
Das Fenster zeigt unbekannte Namen links; sie können per Drag & Drop einer bekannten Liste zugeordnet werden. Über **Export** werden die Zuordnungen ausgegeben.
