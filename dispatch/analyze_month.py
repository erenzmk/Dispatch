from __future__ import annotations

import argparse
import csv
from pathlib import Path
from typing import Iterable

from .process_reports import load_calls, safe_load_workbook
from .name_aliases import load_aliases


def _read_names_from_liste(liste: Path, sheet: str) -> list[str]:
    """Return technician names from column A of *sheet* in *liste*."""
    wb = safe_load_workbook(liste, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise KeyError(f"Worksheet {sheet} does not exist in {liste}")
        ws = wb[sheet]
        return [
            str(ws.cell(row=r, column=1).value).strip()
            for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=1).value
        ]
    finally:
        wb.close()


def analyze_month(month_dir: Path, liste: Path, output: Path) -> None:
    """Analyse all day folders inside *month_dir* and write summary to CSV."""
    load_aliases(liste)
    month_sheet = month_dir.name
    valid_names = _read_names_from_liste(liste, month_sheet)
    expected = set(valid_names)
    found: set[str] = set()

    for day in sorted(p for p in month_dir.iterdir() if p.is_dir()):
        morning_files = list(day.glob("*7*.xlsx"))
        if morning_files:
            _, morning = load_calls(morning_files[0], valid_names)
            found.update(morning)
        evening_files = list(day.glob("*19*.xlsx"))
        if evening_files:
            _, evening = load_calls(evening_files[0], valid_names)
            found.update(evening)

    missing_calls = sorted(expected - found)
    regional_mismatch = sorted(found - expected)

    with output.open("w", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["category", "technician"])
        for name in missing_calls:
            writer.writerow(["no_calls", name])
        for name in regional_mismatch:
            writer.writerow(["region_mismatch", name])


def main(argv: Iterable[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Analyse monthly reports")
    parser.add_argument("month_dir", type=Path, help="Directory containing day folders")
    parser.add_argument("liste", type=Path, help="Path to Liste.xlsx")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("analysis.csv"),
        help="Output CSV file",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)
    analyze_month(args.month_dir, args.liste, args.output)


if __name__ == "__main__":
    main()
