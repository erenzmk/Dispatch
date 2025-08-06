from __future__ import annotations

"""Grafische Oberfläche für die Auswertung von Dispatch-Reports.

Dieses Skript bietet eine GUI zur Auswahl eines Tages oder eines gesamten
Monats und führt anschließend die bekannten Auswertungen durch. Die
Funktionen können auch ohne GUI verwendet werden, was das Testen in
Headless-Umgebungen erleichtert.
"""

from datetime import datetime
from pathlib import Path
import subprocess

# Verzeichnisse für Logs und Ergebnisse
LOG_DIR = Path("logs")
RESULTS_DIR = Path("results")
PROTOCOL_FILE = LOG_DIR / "arbeitsprotokoll.md"


def _log(message: str) -> None:
    """Schreibt eine Meldung in das Arbeitsprotokoll."""
    LOG_DIR.mkdir(exist_ok=True)
    with PROTOCOL_FILE.open("a", encoding="utf-8") as fh:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        fh.write(f"{timestamp} - {message}\n")


def summarize_day(day_dir: Path, liste: Path) -> None:
    """Fasst alle Reports eines Tages nach Techniker-ID zusammen."""
    RESULTS_DIR.mkdir(exist_ok=True)
    for excel in sorted(day_dir.glob("*.xlsx")):
        output = RESULTS_DIR / f"{day_dir.name}_{excel.stem}_summary.csv"
        subprocess.run(
            [
                "python",
                "-m",
                "dispatch.main",
                "summarize-id",
                str(excel),
                str(liste),
                "--output",
                str(output),
            ],
            check=True,
        )
        _log(f'Report "{excel}" -> "{output}"')


def process_month(month_dir: Path, liste: Path, output: Path) -> None:
    """Verarbeitet einen kompletten Monat und erstellt Tageszusammenfassungen."""
    LOG_DIR.mkdir(exist_ok=True)
    subprocess.run(
        [
            "python",
            "-m",
            "dispatch.main",
            "run-all",
            str(month_dir),
            str(liste),
            "--output",
            str(output),
        ],
        check=True,
    )
    for day_dir in sorted(p for p in month_dir.iterdir() if p.is_dir()):
        summarize_day(day_dir, liste)
    _log(f'run_all_gui.py ausgeführt mit "{month_dir}" "{liste}" "{output}"')


def run_gui() -> None:
    """Startet die grafische Oberfläche."""
    import PySimpleGUI as sg

    sg.theme("SystemDefault")

    layout = [
        [sg.Text("Reports-Verzeichnis"), sg.Input("data/reports", key="-ROOT-"), sg.FolderBrowse("Wählen")],
        [sg.Text("Technikerliste"), sg.Input("Liste.xlsx", key="-LISTE-"), sg.FileBrowse("Wählen")],
        [
            sg.Text("Datum"),
            sg.Input(key="-DATE-"),
            sg.CalendarButton("Datum wählen", target="-DATE-", format="%Y-%m-%d"),
        ],
        [sg.Button("Gesamten Monat wählen", key="-MONTH-"), sg.Text("Modus: Tag", key="-MODE-")],
        [sg.Button("Start"), sg.Button("Beenden")],
    ]

    window = sg.Window("Dispatch Auswertung", layout)
    month_mode = False

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Beenden"):
            break
        if event == "-MONTH-":
            month_mode = True
            window["-MODE-"].update("Modus: Monat")
        if event == "Start":
            root = Path(values["-ROOT-"])
            liste = Path(values["-LISTE-"])
            date_str = values.get("-DATE-")
            if not date_str:
                sg.popup_error("Bitte ein Datum auswählen.")
                continue
            date = datetime.strptime(date_str, "%Y-%m-%d").date()
            month_dir = root / date.strftime("%Y-%m")
            if month_mode:
                process_month(month_dir, liste, Path("report.csv"))
            else:
                day_dir = month_dir / date.strftime("%d")
                summarize_day(day_dir, liste)
                _log(f'run_all_gui.py ausgeführt mit "{day_dir}" "{liste}"')
            sg.popup("Fertig.")
            month_mode = False
            window["-MODE-"].update("Modus: Tag")

    window.close()


if __name__ == "__main__":
    run_gui()
