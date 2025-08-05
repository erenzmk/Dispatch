import argparse
import datetime as dt
from pathlib import Path
from typing import Union
import pandas as pd
from dispatch.name_aliases import canonical_name


def process_report(file_path: Union[str, Path], technician_name: str) -> pd.DataFrame:
    """Lese einen Excel-Bericht ein und klassifiziere Calls."""

    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Berichtsdatei nicht gefunden: {file_path}")

    # Alle Reiter laden und zu einem DataFrame kombinieren
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    df = pd.concat(all_sheets.values(), ignore_index=True)

    # Prüfen, ob alle benötigten Spalten vorhanden sind
    expected_cols = ["Techniker", "Callnr", "Erstellt"]
    missing = [col for col in expected_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Fehlende Spalten im Bericht: {missing}")

    # Nur Callnummern behalten, die mit '17' beginnen
    df["Callnr"] = df["Callnr"].astype(str)
    df = df[df["Callnr"].str.startswith("17")]

    # Namen normalisieren und auf den angegebenen Techniker filtern
    valid_names: set[str] = set(df["Techniker"].astype(str)) | {technician_name}

    def _norm(name: str) -> str:
        return canonical_name(name.strip(), valid_names).strip().lower()

    df["__techniker_norm"] = df["Techniker"].astype(str).map(_norm)
    target = _norm(technician_name)
    df = df[df["__techniker_norm"] == target].drop(columns="__techniker_norm")

    # Erstellungsdatum parsen (deutsches Format Tag.Monat.Jahr)
    df["Erstellt"] = pd.to_datetime(df["Erstellt"], dayfirst=True)

    # Heutiges Datum bestimmen und Status setzen
    today = dt.datetime.now().date()
    df["Status"] = df["Erstellt"].dt.date.apply(lambda d: "neu" if d == today else "alt")

    return df


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Techniker-Calls aus Berichten auslesen und klassifizieren"
    )
    parser.add_argument("report_file", help="Pfad zur Excel-Datei (z. B. '7 Uhr.xlsx')")
    parser.add_argument(
        "--techniker",
        default="Ahmad, Daniyal (Keskin)",
        help="Name des Technikers genau wie im Bericht",
    )
    parser.add_argument(
        "--output", default="output_calls.xlsx", help="Pfad für die Ausgabedatei"
    )
    args = parser.parse_args()

    try:
        result_df = process_report(args.report_file, args.techniker)
    except FileNotFoundError as err:
        print(err)
        return

    result_df.to_excel(args.output, index=False)
    print(f"Ergebnis gespeichert: {args.output}")


if __name__ == "__main__":
    main()
