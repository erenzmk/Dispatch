"""
Automate dispatch call summaries.

This script processes daily technician call reports. For each day directory
(e.g. ``Juli_25/01.07``) it expects a morning file named with ``7`` and
extracts calls per technician. Calls are categorised into
``total``, ``new`` (opened on the previous business day) and ``old``
(otherwise).

The aggregated values are written into ``Liste.xlsx`` on the sheet for the
corresponding month. The workbook contains weekly column blocks consisting of
The aggregated values are written into ``Liste.xlsx`` on the sheet matching the
name of the month directory (e.g. ``Juli_25``). The workbook contains weekly
column blocks consisting of
13 columns:

    name, date, weekday, pudo, pickup time, valid, info, pre-closed,
    total calls, old calls, new calls, details, mails

Blocks are repeated for each week and separated by an empty column.  The first
block starts at column ``A``.  To update values for a given date ``d`` the
block index is ``(d.day-1)//7`` and the starting column is ``1 + index*14``.

The script requires :mod:`openpyxl` for reading and writing Excel files.
"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Dict

try:
    from openpyxl import load_workbook
except Exception as exc:  # pragma: no cover - missing dependency
    raise SystemExit("openpyxl is required: {}".format(exc))


@@ -114,34 +115,34 @@ def update_liste(liste: Path, month_sheet: str, day: dt.date,
    start_col = 1 + week_index * 14

    for row in range(2, ws.max_row + 1):
        name_cell = ws.cell(row=row, column=1)
        tech = str(name_cell.value).strip() if name_cell.value else None
        if not tech or tech not in summary:
            continue
        data = summary[tech]
        ws.cell(row=row, column=start_col + 1).value = day.toordinal() + 693594
        ws.cell(row=row, column=start_col + 2).value = PREV_DAY_MAP[day.weekday()]
        ws.cell(row=row, column=start_col + 8).value = data["total"]
        ws.cell(row=row, column=start_col + 9).value = data["old"]
        ws.cell(row=row, column=start_col + 10).value = data["new"]
    wb.save(liste)


def main():
    parser = argparse.ArgumentParser(description="Process dispatch reports")
    parser.add_argument("day_dir", type=Path,
                        help="Directory containing daily reports (e.g. Juli_25/01.07)")
    parser.add_argument("liste", type=Path, help="Path to Liste.xlsx")
    args = parser.parse_args()

    day_str = f"{args.day_dir.name}.2025"
    day = dt.datetime.strptime(day_str, "%d.%m.%Y").date()
    month_sheet = day.strftime("%B_%y").capitalize()
    month_sheet = args.day_dir.parent.name

    morning = next(args.day_dir.glob("*7*.xlsx"))
    target_date, summary = load_calls(morning)
    update_liste(args.liste, month_sheet, target_date, summary)


if __name__ == "__main__":
    main()
