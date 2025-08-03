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

from process_reports import load_calls, safe_load_workbook


def gather_valid_names(liste: Path) -> list[str]:
    """Return a list of technician names from ``Liste.xlsx``.

    Only the first column is inspected which is sufficient for our matching
    needs.  Empty cells are ignored.
    """

    names: list[str] = []
    with closing(safe_load_workbook(liste, read_only=True)) as wb:
        ws = wb.active
        for cell in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            value = cell[0]
            if value:
                names.append(str(value).strip())
    return names


def aggregate_warnings(report_dir: Path, valid_names: list[str]) -> Counter[str]:
    """Process all Excel files below *report_dir* and count unknown names."""

    counter: Counter[str] = Counter()
    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    root_logger = logging.getLogger()
    root_logger.addHandler(handler)
    previous_level = root_logger.level
    root_logger.setLevel(logging.WARNING)
    try:
        for file in sorted(report_dir.rglob("*.xlsx")):
            stream.seek(0)
            stream.truncate(0)
            load_calls(file, valid_names)
            stream.seek(0)
            for line in stream.read().splitlines():
                match = re.search(r"'([^']+)'", line)
                if match:
                    counter[match.group(1)] += 1
    finally:
        root_logger.removeHandler(handler)
        handler.close()
        root_logger.setLevel(previous_level)
    return counter


def main() -> None:  # pragma: no cover - convenience script
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report_dir", type=Path, help="Path to Juli_25 directory")
    parser.add_argument(
        "--liste", type=Path, default=Path("Liste.xlsx"), help="Path to Liste.xlsx"
    )
    args = parser.parse_args()

    valid = gather_valid_names(args.liste)
    counter = aggregate_warnings(args.report_dir, valid)
    for name, count in counter.most_common():
        print(f"{name}: {count}")


if __name__ == "__main__":  # pragma: no cover - script entry point
    main()

