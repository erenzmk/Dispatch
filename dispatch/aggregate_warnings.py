"""Aggregate unknown technician warnings across July reports.

This optional helper walks through a directory containing the July reports
(``Juli_25``) and feeds all Excel files into :func:`process_reports.load_calls`.
Any warnings about technicians that could not be matched are collected and
counted so dispatchers can spot missing aliases.

Usage::

    python aggregate_warnings.py Juli_25 --liste Liste.xlsx

The script prints a simple summary mapping the unresolved technician names to
the number of occurrences in the provided reports.
"""

from __future__ import annotations

import argparse
import io
import logging
import re
from collections import Counter
from contextlib import closing
from pathlib import Path
from typing import Iterable

from . import process_reports
from .process_reports import load_calls, safe_load_workbook


logger = logging.getLogger(__name__)


def gather_valid_names(liste: Path, sheet_name: str | None = None) -> list[str]:
    """Gib eine sortierte Liste einzigartiger Techniker aus ``Liste.xlsx`` zurück.

    Ohne Angabe von *sheet_name* wird das erste Tabellenblatt gewählt, dessen
    Titel ``"Technikernamen"`` enthält. Ist kein entsprechendes Blatt
    vorhanden, wird eine :class:`ValueError` mit den vorhandenen Blattnamen
    ausgelöst. Die Spalten ``Technikername`` und ``PUOOS`` werden eingelesen,
    leere Zellen ignoriert und doppelte Einträge entfernt.
    """

    names: set[str] = set()
    with closing(safe_load_workbook(liste, read_only=True)) as wb:
        if sheet_name is None:
            target = next(
                (name for name in wb.sheetnames if "technikernamen" in name.lower()),
                None,
            )
            if target is None:
                raise ValueError(
                    f"Kein Tabellenblatt mit 'Technikernamen' in {liste}; vorhanden: {', '.join(wb.sheetnames)}"
                )
            ws = wb[target]
        else:
            if sheet_name not in wb.sheetnames:
                raise ValueError(
                    f"Tabellenblatt {sheet_name!r} fehlt in {liste}; vorhanden: {', '.join(wb.sheetnames)}"
                )
            ws = wb[sheet_name]

        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        wanted = [
            idx
            for idx, title in enumerate(header)
            if str(title).strip().lower() in {"technikername", "puoos"}
        ]
        if not wanted:
            wanted = list(range(len(header)))

        for row in ws.iter_rows(min_row=2, values_only=True):
            for idx in wanted:
                if idx < len(row):
                    value = row[idx]
                    if value:
                        names.add(str(value).strip())
    return sorted(names)


def aggregate_warnings(report_dir: Path, valid_names: list[str]) -> Counter[str]:
    """Process all Excel files below *report_dir* and count unknown names."""

    counter: Counter[str] = Counter()
    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    handler.setLevel(logging.WARNING)

    capture_logger = logger.getChild("process")
    capture_logger.propagate = False
    capture_logger.addHandler(handler)

    original_logger = process_reports.logger
    process_reports.logger = capture_logger
    try:
        for file in sorted(report_dir.rglob("*.xlsx")):
            if file.name.lower().startswith("liste"):
                continue  # skip the aggregated Liste workbook
            stream.seek(0)
            stream.truncate(0)
            try:
                load_calls(file, valid_names)
            except ValueError as exc:
                logger.error("%s: %s", file, exc)
                continue
            stream.seek(0)
            for line in stream.read().splitlines():
                match = re.search(r"'([^']+)'", line)
                if match:
                    counter[match.group(1)] += 1
    finally:
        capture_logger.removeHandler(handler)
        handler.close()
        process_reports.logger = original_logger
    return counter


def main(argv: Iterable[str] | None = None) -> None:  # pragma: no cover - convenience script
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report_dir", type=Path, help="Path to Juli_25 directory")
    parser.add_argument(
        "--liste", type=Path, default=Path("Liste.xlsx"), help="Path to Liste.xlsx"
    )
    parser.add_argument("--sheet", help="Name des Tabellenblatts mit Technikern")
    args = parser.parse_args(list(argv) if argv is not None else None)

    valid = gather_valid_names(args.liste, sheet_name=args.sheet)
    counter = aggregate_warnings(args.report_dir, valid)
    for name, count in counter.most_common():
        print(f"{name}: {count}")


if __name__ == "__main__":  # pragma: no cover - script entry point
    main()

