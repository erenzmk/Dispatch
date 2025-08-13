"""Parser für die OR Liste im 2-Spalten-Muster.

Dieses Modul liest eine Excel-Datei, in der Techniker und Datum in
paarweise angeordneten Spalten stehen (Name|Datum). Die Werte werden
zugeordnet und nach Datum gruppiert.
"""
from __future__ import annotations

from pathlib import Path
import pandas as pd
import unicodedata

# Basisverzeichnis des Projekts: zwei Ebenen über dieser Datei
BASE_DIR = Path(__file__).resolve().parent.parent
XLSX_PATH = BASE_DIR / "data" / "Liste.xlsx"


__all__ = ["parse_or_liste", "group_by_day"]


def canon_name(s: str) -> str:
    """Normalisiere *s* und führe bekannte Alias-Korrekturen durch."""
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s_norm = unicodedata.normalize("NFKD", s)
    s_ascii = "".join(ch for ch in s_norm if not unicodedata.combining(ch))
    fixes = {
        "Kuersad": "Kürşad",
        "Serghei S.": "Serghei S.",
    }
    return fixes.get(s_ascii, s_ascii)


def parse_or_liste(
    xlsx_path: Path = XLSX_PATH, sheet_name: str | int = 0
) -> pd.DataFrame:
    """Lese die OR Liste und liefere eine DataFrame mit ``date`` und ``tech``.

    *xlsx_path* zeigt auf ``data/Liste.xlsx`` im Projekt. Enthält die
    Arbeitsmappe mehrere Monatsblätter, kann optional *sheet_name*
    angegeben werden; fehlt es, wird das erste Blatt genutzt.
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)
    records: list[dict[str, object]] = []

    # Spalten anhand der Kopfzeile identifizieren: jede Spalte mit
    # "name" markiert den Beginn eines (Name|Datum)-Blocks.
    header = df.iloc[0].astype(str).str.strip().str.lower()
    name_cols = [idx for idx, val in enumerate(header) if val == "name"]

    for name_col in name_cols:
        date_col = name_col + 1
        if date_col >= df.shape[1]:
            continue

        sample_dates = pd.to_datetime(
            df.iloc[1:, date_col], errors="coerce", format="mixed"
        )
        if sample_dates.notna().sum() == 0:
            continue

        for r in range(1, df.shape[0]):
            raw_name = df.iat[r, name_col]
            raw_date = df.iat[r, date_col]

            if pd.isna(raw_name) and pd.isna(raw_date):
                continue

            name = canon_name(str(raw_name)) if pd.notna(raw_name) else ""
            date = pd.to_datetime(raw_date, errors="coerce", format="mixed")

            if name and pd.notna(date):
                records.append({"date": date.normalize(), "tech": name})

    if not records:
        return pd.DataFrame(columns=["date", "tech"])

    out = (
        pd.DataFrame.from_records(records)
        .drop_duplicates()
        .sort_values(["date", "tech"])
        .reset_index(drop=True)
    )
    return out


def group_by_day(df_assign: pd.DataFrame) -> dict[pd.Timestamp, list[str]]:
    """Gruppiere die Zuordnungen nach Datum."""
    grouped = df_assign.groupby("date")["tech"].apply(list).to_dict()
    return grouped
