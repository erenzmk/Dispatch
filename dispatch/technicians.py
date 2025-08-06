from __future__ import annotations

from pathlib import Path
from openpyxl import load_workbook


def load_id_map(liste_path: Path) -> dict[str, str]:
    """Lese ein Mapping von Techniker-IDs zu Namen aus einer Excel-Datei.

    Erwartet eine Tabelle mit den Spalten ``ID`` und ``Techniker`` im ersten
    Arbeitsblatt. Whitespace wird entfernt und leere Zeilen werden ignoriert.
    """
    wb = load_workbook(liste_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
    id_col = headers.get("ID")
    tech_col = headers.get("Techniker")
    result: dict[str, str] = {}

    if id_col is None or tech_col is None:
        wb.close()
        return result

    for row in ws.iter_rows(min_row=2, values_only=True):
        id_value = row[id_col - 1]
        tech_value = row[tech_col - 1]
        if id_value is None or tech_value is None:
            continue
        key = str(id_value).strip()
        value = str(tech_value).strip()
        if key and value:
            result[key] = value

    wb.close()
    return result
