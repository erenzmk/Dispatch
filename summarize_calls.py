import argparse
from pathlib import Path
from typing import Iterable
import pandas as pd

MAPPING_SHEETS = ["Techniker DK", "Berlin dk", "Berlin rest"]


def load_technicians(xls: pd.ExcelFile, mapping_sheets: Iterable[str]) -> set[str]:
    """Alle Techniker-Namen aus den Zuordnungsblättern sammeln."""
    names: set[str] = set()
    for sheet in mapping_sheets:
        if sheet not in xls.sheet_names:
            continue
        df = pd.read_excel(xls, sheet_name=sheet)
        first = df.get("first", pd.Series(dtype=str)).fillna("")
        last = df.get("last", pd.Series(dtype=str)).fillna("")
        full = (first + " " + last).str.strip()
        names.update(full[full != ""])  # Leere Einträge ignorieren
    return names


def summarize_report(file_path: Path) -> pd.DataFrame:
    """Bericht einlesen und Calls pro Techniker zusammenfassen."""
    xls = pd.ExcelFile(file_path)
    technicians = load_technicians(xls, MAPPING_SHEETS)
    if not technicians:
        raise ValueError("Keine Techniker in den Mapping-Tabellen gefunden")

    # Basisdatum aus Zelle A2 des ersten Datenblatts
    data_sheets = [s for s in xls.sheet_names if s not in MAPPING_SHEETS]
    if not data_sheets:
        raise ValueError("Keine Datenblätter gefunden")
    tmp = pd.read_excel(xls, sheet_name=data_sheets[0], header=None)
    file_date = pd.to_datetime(tmp.iloc[1, 0], dayfirst=True).date()

    records = []
    for sheet in data_sheets:
        df = pd.read_excel(xls, sheet_name=sheet)
        for _, row in df.iterrows():
            name = str(row.iloc[0]).strip()
            if name not in technicians:
                continue
            call = str(row.iloc[2]).strip()
            if not call.startswith("17"):
                continue
            h_val = row.iloc[7]
            if pd.isnull(h_val):
                continue
            call_date = pd.to_datetime(h_val, dayfirst=True).date()
            bd_range = pd.bdate_range(call_date, file_date)
            diff = len(bd_range) - 1
            status = "new" if diff <= 1 else "old"
            records.append({"technician": name, "status": status})

    if not records:
        return pd.DataFrame(columns=["technician", "date", "new", "old", "total"])

    df_res = pd.DataFrame(records)
    summary = (
        df_res.groupby(["technician", "status"]).size().unstack(fill_value=0)
    )
    summary = summary.reindex(columns=["new", "old"], fill_value=0)
    summary["total"] = summary.sum(axis=1)
    summary = summary.reset_index()
    summary["date"] = file_date
    return summary[["technician", "date", "new", "old", "total"]]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Zähle neue und alte Calls je Techniker aus einem Bericht"
    )
    parser.add_argument("excel_file", type=Path, help="Pfad zur Excel-Datei")
    parser.add_argument(
        "--output", type=Path, help="Pfad für die Ausgabedatei (CSV)")
    args = parser.parse_args()

    summary = summarize_report(args.excel_file)
    if args.output:
        summary.to_csv(args.output, index=False)
        print(f"Ergebnis gespeichert in {args.output}")
    else:
        print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
