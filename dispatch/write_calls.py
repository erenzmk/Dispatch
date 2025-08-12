from pathlib import Path
from typing import Mapping
import pandas as pd
from openpyxl import load_workbook

# Zuordnung Woche -> (Startspalte Datum, Startspalte Call)
WEEK_START_COLUMNS: Mapping[int, tuple[str, str]] = {
    1: ("A", "B"),
    2: ("O", "P"),
    3: ("AC", "AD"),
}


def write_calls(workbook: Path | str, records: pd.DataFrame, sheet_name: str = "Juli") -> None:
    """Trage Call-Werte anhand von Namen und Datum in das Monatsblatt ein.

    ``records`` muss die Spalten ``name``, ``date`` und ``value`` enthalten.
    """
    wb = load_workbook(workbook)
    ws = wb[sheet_name]

    # Namen des ersten Blocks auslesen und auf Indizes abbilden
    name_map: dict[str, int] = {}
    for idx, row in enumerate(ws["A2":"A27"], start=0):
        cell = row[0]
        if cell.value:
            name_map[str(cell.value).strip()] = idx

    for _, rec in records.iterrows():
        name = str(rec["name"]).strip()
        date = pd.to_datetime(rec["date"]).date()
        value = rec["value"]

        if name not in name_map:
            continue
        week = ((date.day - 1) // 7) + 1
        week_cols = WEEK_START_COLUMNS.get(week)
        if not week_cols:
            continue
        _, call_col = week_cols
        day_offset = date.weekday()
        row = 2 + name_map[name] + 26 * day_offset
        ws[f"{call_col}{row}"] = value

    wb.save(workbook)
