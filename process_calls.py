import argparse
import datetime as dt
from pathlib import Path
from typing import Union
import pandas as pd


def vorheriger_werktag(referenz: dt.date) -> dt.date:
    """Gib den letzten Werktag vor ``referenz`` zurück."""

    tag = referenz - dt.timedelta(days=1)
    while tag.weekday() >= 5:  # 5 = Samstag, 6 = Sonntag
        tag -= dt.timedelta(days=1)
    return tag


def process_report(file_path: Union[str, Path], technician_name: str) -> pd.DataFrame:
    """Lese einen Excel-Bericht ein und klassifiziere Calls."""

    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Berichtsdatei nicht gefunden: {file_path}")

    # Alle Reiter laden
    all_sheets = pd.read_excel(file_path, sheet_name=None)

    report_date = None
    call_frames = []
    for sheet in all_sheets.values():
        if {"Techniker", "Callnr", "Erstellt"}.issubset(sheet.columns):
            call_frames.append(sheet)
        if report_date is None and "Berichtstag" in sheet.columns:
            value = sheet["Berichtstag"].dropna().iloc[0]
            report_date = pd.to_datetime(value, dayfirst=True).date()

    if not call_frames:
        raise ValueError("Keine gültigen Datenblätter gefunden")

    df = pd.concat(call_frames, ignore_index=True)

    # Prüfen, ob alle benötigten Spalten vorhanden sind
    expected_cols = ["Techniker", "Callnr", "Erstellt"]
    missing = [col for col in expected_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Fehlende Spalten im Bericht: {missing}")

    # Nur Callnummern behalten, die mit '17' beginnen
    df["Callnr"] = df["Callnr"].astype(str)
    df = df[df["Callnr"].str.startswith("17")]

    # Auf den angegebenen Techniker filtern
    df = df[df["Techniker"] == technician_name]

    # Erstellungsdatum parsen (deutsches Format Tag.Monat.Jahr)
    df["Erstellt"] = pd.to_datetime(df["Erstellt"], dayfirst=True)

    # Referenzdatum bestimmen und Status setzen
    if report_date is None:
        report_date = dt.date.today()
    werk_vortag = vorheriger_werktag(report_date)
    df["Status"] = df["Erstellt"].dt.date.apply(
        lambda d: "neu" if d == werk_vortag else "alt"
    )

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
