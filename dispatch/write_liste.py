"""Schreibe Tagesdaten in die Projektdatei ``Liste.xlsx``.

Dieses Skript liest tägliche Reports aus ``data/reports`` ein, fasst die
Daten pro Techniker zusammen und trägt sie in das Monatsblatt ``Juli_25``
der Projektdatei ein. Vorhandene Blöcke werden wiederverwendet, neue
Blöcke bei Bedarf rechts angefügt. Es werden ausschließlich Zellwerte
geändert, Formatierungen bleiben erhalten.
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
from pathlib import Path
from typing import Dict, List

import pandas as pd
from openpyxl import load_workbook

from .name_aliases import canonical_name, ALIASES, refresh_alias_map

INFILE = Path(r"C:\Users\egencer\Documents\GitHub\Dispatch\data\Liste.xlsx")
BLOCK_FIELDS = [
    "name",
    "date",
    "weekday",
    "pudo",
    "pickup time",
    "valid",
    "info",
    "pre-closed",
    "total Calls",
    "old calls",
    "new Calls",
    "details",
    "Mails",
]
BLOCK_WIDTH = len(BLOCK_FIELDS) + 1  # +1 für Trenner

logger = logging.getLogger("write_liste")


def load_mapping(wb_path: Path) -> tuple[Dict[str, str], List[str]]:
    """Lese Alias-Mapping und Technikerreihenfolge.

    Die Tabelle ``Technikernamen + PUDO`` enthält die Technikerliste. Aus den
    Spalten ``first`` und ``last`` wird der kanonische Name aufgebaut. Die
    Spalte ``dk`` liefert bekannte Kurzformen, die auf den kanonischen Namen
    abgebildet werden.
    """

    df = pd.read_excel(wb_path, sheet_name="Technikernamen + PUDO", header=1)
    df = df[["first", "last", "dk"]].dropna(subset=["first"])
    df["name"] = (df["first"].fillna("") + " " + df["last"].fillna("")).str.strip()

    mapping: Dict[str, str] = {}
    for _, row in df.iterrows():
        alias = str(row.get("dk", "")).strip()
        name = row["name"]
        if alias and name:
            mapping[alias] = name

    tech_order = df["name"].tolist()
    return mapping, tech_order


def detect_blocks(ws) -> List[int]:
    """Finde Startspalten aller Tagesblöcke anhand der Kopfzeile."""

    headers = [str(cell.value).strip().lower() if cell.value else "" for cell in ws[1]]
    starts: List[int] = []
    for idx, value in enumerate(headers, start=1):
        if value == "name":
            match = True
            for off, field in enumerate(BLOCK_FIELDS):
                header = headers[idx - 1 + off] if idx - 1 + off < len(headers) else ""
                if header != field.lower():
                    match = False
                    break
            if match:
                starts.append(idx)
    return starts


def find_or_create_day_block(ws, target_date: dt.date) -> tuple[int, bool]:
    """Finde den Block für ``target_date`` oder lege einen neuen an.

    Rückgabe ist ``(startspalte, neu_angelegt)``.
    """

    starts = detect_blocks(ws)
    first = starts[0]
    for start in starts:
        cell = ws.cell(row=2, column=start + 1).value
        if isinstance(cell, dt.datetime):
            cell = cell.date()
        if cell == target_date:
            return start, False

    new_start = starts[-1] + BLOCK_WIDTH
    for off in range(len(BLOCK_FIELDS)):
        ws.cell(row=1, column=new_start + off).value = ws.cell(row=1, column=first + off).value
    row_idx = 2
    while True:
        src = ws.cell(row=row_idx, column=first)
        if not src.value:
            break
        ws.cell(row=row_idx, column=new_start).value = src.value
        row_idx += 1
    return new_start, True


def aggregate_rows(rows: pd.DataFrame) -> dict[str, object]:
    """Fasse mehrere Zeilen nach den vorgegebenen Regeln zusammen."""

    result: dict[str, object] = {}
    numeric = ["pre-closed", "total Calls", "old calls", "new Calls", "Mails"]
    text = ["info", "details", "pudo"]

    for col in numeric:
        vals = pd.to_numeric(rows[col], errors="coerce")
        total = vals.sum(min_count=1)
        if pd.notna(total):
            result[col] = float(total)
    for col in text:
        parts = [str(v).strip() for v in rows[col].dropna() if str(v).strip()]
        if parts:
            result[col] = " | ".join(sorted(set(parts)))
    times = pd.to_datetime(rows["pickup time"], errors="coerce")
    if not times.isna().all():
        result["pickup time"] = times.min().time()
    valids = rows["valid"].dropna().astype(bool)
    if not valids.empty:
        result["valid"] = bool(valids.any())
    return result


def write_day(
    ws,
    mapping: Dict[str, str],
    tech_order: List[str],
    day_df: pd.DataFrame,
    date: dt.date,
    normalize: bool = True,
) -> None:
    """Schreibe alle Daten des Tages ``date`` in das Arbeitsblatt."""

    alias_map = {**mapping}
    ALIASES.update(alias_map)
    refresh_alias_map()

    normalized = 0
    skipped = 0

    def norm_name(name: str) -> str:
        name = str(name).strip()
        return alias_map.get(name, canonical_name(name, tech_order))

    if "date" not in day_df.columns:
        day_df["date"] = date
    for idx, row in day_df.iterrows():
        if pd.isna(row.get("date")):
            day_df.at[idx, "date"] = date
            normalized += 1
            continue
        rdate = pd.to_datetime(row["date"]).date()
        if rdate != date:
            if normalize:
                day_df.at[idx, "date"] = date
                normalized += 1
            else:
                day_df.at[idx, "_skip"] = True
                skipped += 1
    if "_skip" in day_df.columns:
        day_df = day_df[~day_df["_skip"]]

    day_df["name"] = day_df["name"].map(norm_name)
    grouped = day_df.groupby(["name", "date"], dropna=True)
    aggregated: Dict[str, dict[str, object]] = {}
    for (name, _), rows in grouped:
        aggregated[name] = aggregate_rows(rows)

    start_col, created = find_or_create_day_block(ws, date)
    logger.info(
        "Block für %s %s", date.isoformat(), "angelegt" if created else "gefunden"
    )

    for tech, values in aggregated.items():
        if tech not in tech_order:
            continue
        row = tech_order.index(tech) + 2
        # Date and weekday setzen
        ws.cell(row=row, column=start_col + 1, value=date)
        ws.cell(row=row, column=start_col + 2, value=date.strftime("%A"))
        for off, field in enumerate(BLOCK_FIELDS):
            if field in ("date", "weekday", "name"):
                continue
            value = values.get(field)
            if value in (None, ""):
                continue
            if isinstance(value, float) and pd.isna(value):
                continue
            col = start_col + BLOCK_FIELDS.index(field)
            existing = ws.cell(row=row, column=col).value
            if existing in (None, ""):
                ws.cell(row=row, column=col, value=value)
        ws.cell(row=row, column=start_col, value=tech)

    logger.info(
        "Normalisierte Zeilen: %d, übersprungene Zeilen: %d", normalized, skipped
    )


def collect_day_df(day_dir: Path) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for file in sorted(day_dir.iterdir()):
        if file.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            frames.append(pd.read_excel(file))
        elif file.suffix.lower() in {".csv", ".txt"}:
            frames.append(pd.read_csv(file, sep=";"))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--infile", default=str(INFILE))
    ap.add_argument("--out", default=None)
    ap.add_argument("--month", required=True)
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--normalize-date", choices=["true", "false"], default="true")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    wb = load_workbook(args.infile)
    ws = wb["Juli_25"]

    mapping, tech_order = load_mapping(Path(args.infile))

    year, month = map(int, args.month.split("-"))
    days = (dt.date(year, month, 28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    for day in range(1, days.day + 1):
        date = dt.date(year, month, day)
        day_dir = Path("data") / "reports" / f"{year:04d}-{month:02d}" / f"{day:02d}"
        day_df = collect_day_df(day_dir)
        if day_df.empty:
            continue
        write_day(ws, mapping, tech_order, day_df, date, normalize=args.normalize_date == "true")

    out_file = args.infile if args.inplace or not args.out else args.out
    wb.save(out_file)


if __name__ == "__main__":
    main()
