"""Erstelle automatisch den Tagesordner unter ``data/reports``.

Dieses Hilfsskript legt für das aktuelle Datum einen Ordner nach dem Schema
``data/reports/YYYY-MM/TT`` an und gibt den Pfad auf der Konsole aus.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path


def create_day_dir(base: Path = Path("data/reports")) -> Path:
    """Erzeuge den Ordner für das heutige Datum und gib den Pfad zurück."""
    today = dt.date.today()
    month_dir = base / today.strftime("%Y-%m")
    day_dir = month_dir / today.strftime("%d")
    day_dir.mkdir(parents=True, exist_ok=True)
    return day_dir


if __name__ == "__main__":
    path = create_day_dir()
    print(path)
