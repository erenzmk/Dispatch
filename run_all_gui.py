from __future__ import annotations

"""Grafische Oberfläche für die Auswertung von Dispatch-Reports.

Dieses Skript bietet eine GUI zur Auswahl eines Tages oder eines gesamten
Monats und führt anschließend die bekannten Auswertungen durch. Die
Funktionen können auch ohne GUI verwendet werden, was das Testen in
Headless-Umgebungen erleichtert.

Hinweis: Die GUI basiert auf PySimpleGUI 5.x.
"""

from datetime import datetime
from pathlib import Path
import subprocess
import sys

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


def _popup_error(message: str) -> None:
    """Zeigt eine Fehlermeldung in der GUI an, falls möglich."""
    try:
        import PySimpleGUI as sg  # type: ignore
    except Exception:
        return
    sg.popup_error(message)


def summarize_day(day_dir: Path, liste: Path) -> bool:
    """Fasst alle Reports eines Tages nach Techniker-ID zusammen."""
    RESULTS_DIR.mkdir(exist_ok=True)
    try:
        subprocess.run(
            [
                "python",
                "-m",
                "dispatch.main",
                "process",
                str(day_dir),
                str(liste),
            ],
            check=True,
        )
    except subprocess.CalledProcessError as exc:
        _log(f'Fehler bei dispatch.main process mit "{day_dir}" "{liste}": {exc}')
        _popup_error(f"Fehler bei der Tagesverarbeitung:\n{exc}")
        return False
    else:
        _log(f'dispatch.main process mit "{day_dir}" "{liste}"')

    success = True
    for excel in sorted(day_dir.glob("*.xlsx")):
        output = RESULTS_DIR / f"{day_dir.name}_{excel.stem}_summary.csv"
        try:
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
        except subprocess.CalledProcessError as exc:
            _log(f'Fehler bei Report "{excel}" -> "{output}": {exc}')
            _popup_error(f"Fehler bei Report {excel}:\n{exc}")
            success = False
        else:
            _log(f'Report "{excel}" -> "{output}"')
    return success


def process_month(month_dir: Path, liste: Path, output: Path) -> bool:
    """Verarbeitet einen kompletten Monat und erstellt Tageszusammenfassungen."""
    LOG_DIR.mkdir(exist_ok=True)
    try:
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
    except subprocess.CalledProcessError as exc:
        _log(
            f'Fehler bei run-all mit "{month_dir}" "{liste}" "{output}": {exc}'
        )
        _popup_error(f"Fehler bei der Monatsverarbeitung:\n{exc}")
        return False

    success = True
    for day_dir in sorted(p for p in month_dir.iterdir() if p.is_dir()):
        success &= summarize_day(day_dir, liste)
    if success:
        _log(f'run_all_gui.py ausgeführt mit "{month_dir}" "{liste}" "{output}"')
    return success


def run_gui() -> None:
    """Startet die grafische Oberfläche."""
    import PySimpleGUI as sg
    if not hasattr(sg, "theme"):
        print(
            "PySimpleGUI unterstützt 'theme' nicht. "
            "Bitte Version 5.x über den privaten Index installieren:\n"
            "pip install --extra-index-url https://PySimpleGUI.net/install PySimpleGUI"
        )
        sys.exit(1)

    sg.theme("SystemDefault")

    layout = [
        [sg.Text("Reports-Verzeichnis"), sg.Input("data/reports", key="-ROOT-"), sg.FolderBrowse("Wählen")],
        [sg.Text("Technikerliste"), sg.Input("Liste.xlsx", key="-LISTE-"), sg.FileBrowse("Wählen")],
        [sg.Text("Ausgabedatei"), sg.Input(key="-OUTPUT-"), sg.FileSaveAs("Wählen")],
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
            output_str = values.get("-OUTPUT-")
            if output_str:
                output = Path(output_str)
            else:
                output = Path(f"report_{date.strftime('%Y-%m')}.csv")
            if month_mode:
                ok = process_month(month_dir, liste, output)
            else:
                day_dir = month_dir / date.strftime("%d")
                ok = summarize_day(day_dir, liste)
                if ok:
                    _log(f'run_all_gui.py ausgeführt mit "{day_dir}" "{liste}"')
            if ok:
                sg.popup("Fertig.")
            month_mode = False
            window["-MODE-"].update("Modus: Tag")

    window.close()


if __name__ == "__main__":
    run_gui()
