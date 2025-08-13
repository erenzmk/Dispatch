from __future__ import annotations

"""Hilfsfunktionen für die Auswertung von Dispatch-Reports.

Dieses Modul stellt nur noch die Funktionen ``summarize_day`` und
``process_month`` sowie die Protokollierungsfunktion ``_log`` bereit.
"""

from datetime import datetime
from pathlib import Path
import json
import os
import tempfile
from collections import defaultdict

import pandas as pd

from dispatch import summarize_by_id, process_reports, analyze_month

# Verzeichnisse für Logs und Ergebnisse
_DEFAULT_LOG_DIR = Path(tempfile.gettempdir()) / "dispatch_logs"
LOG_DIR = Path(os.environ.get("DISPATCH_LOG_DIR", str(_DEFAULT_LOG_DIR)))
RESULTS_DIR = Path("results")
PROTOCOL_FILE = LOG_DIR / "arbeitsprotokoll.md"


def _log(message: str, data: object | None = None) -> None:
    """Schreibt eine Meldung in das Arbeitsprotokoll.

    Optionale strukturierte Daten werden formatiert angehängt.
    """
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    with PROTOCOL_FILE.open("a", encoding="utf-8") as fh:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        fh.write(f"{timestamp} - {message}\n")
        if data is not None:
            fh.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")


def summarize_day(
    day_dir: Path,
    liste: Path,
    call_log: dict[str, list[str]] | None = None,
) -> bool:
    """Fasst alle Reports eines Tages nach Techniker-ID zusammen.

    Wird ``call_log`` übergeben, werden die Call-Listen pro Techniker-ID
    abgelegt.
    """

    RESULTS_DIR.mkdir(exist_ok=True)
    try:
        process_reports.main([str(day_dir), str(liste)])
    except Exception as exc:
        _log(f'Fehler bei dispatch.process_reports mit "{day_dir}" "{liste}": {exc}')
        return False
    else:
        _log(f'dispatch.process_reports mit "{day_dir}" "{liste}"')

    success = True
    calls_by_id: defaultdict[str, list[str]] = defaultdict(list)
    for excel in sorted(day_dir.glob("*.xlsx")):
        output = RESULTS_DIR / f"{day_dir.name}_{excel.stem}_summary.csv"
        try:
            summary = summarize_by_id.summarize_report(excel, liste)
            pd.DataFrame(summary).to_csv(output, index=False)
        except Exception as exc:
            _log(f'Fehler bei Report "{excel}" -> "{output}": {exc}')
            success = False
        else:
            _log(f'Report "{excel}" -> "{output}"')
            calls = {row["id"]: row.get("calls", []) for row in summary}
            for tech_id, call_list in calls.items():
                calls_by_id[tech_id].extend(call_list)
    final_calls = {tid: cl for tid, cl in calls_by_id.items()}
    if call_log is not None:
        call_log.update(final_calls)
    if success:
        _log(f'Call-Listen für Tag "{day_dir.name}"', final_calls)
    return success


def process_month(
    month_dir: Path,
    liste: Path,
    output: Path,
    call_log: dict[str, dict[str, list[str]]] | None = None,
) -> bool:
    """Verarbeitet einen kompletten Monat und erstellt Tageszusammenfassungen.

    Wird ``call_log`` übergeben, werden die Call-Listen der einzelnen Tage
    darin gesammelt.
    """

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        process_reports.process_month(month_dir, liste)
        analyze_month.main([str(month_dir), str(liste), "-o", str(output)])
    except Exception as exc:
        _log(
            f'Fehler bei run-all mit "{month_dir}" "{liste}" "{output}": {exc}'
        )
        return False

    success = True
    month_calls: dict[str, dict[str, list[str]]] = {}
    for day_dir in sorted(p for p in month_dir.iterdir() if p.is_dir()):
        day_calls: dict[str, list[str]] = {}
        day_success = summarize_day(day_dir, liste, day_calls)
        month_calls[day_dir.name] = day_calls
        if call_log is not None:
            call_log[day_dir.name] = day_calls
        success &= day_success
    if success:
        _log(
            f'run_dispatch.py ausgeführt mit "{month_dir}" "{liste}" "{output}"'
        )
        _log(f'Call-Listen für Monat "{month_dir.name}"', month_calls)
    return success
