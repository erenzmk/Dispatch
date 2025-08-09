from datetime import date
from pathlib import Path
import sys
from run_all_gui import process_month

LOG_FILE = Path("arbeitsprotokoll.txt")


def log(message: str) -> None:
    with LOG_FILE.open("a", encoding="utf-8") as fh:
        fh.write(f"{date.today().isoformat()} - {message}\n")


def main() -> None:
    month_str = sys.argv[1] if len(sys.argv) > 1 else date.today().strftime("%Y-%m")
    try:
        date.fromisoformat(f"{month_str}-01")
    except ValueError:
        log(f"Ungültiges Monatsformat: {month_str}")
        return
    month_dir = Path("data", "reports", month_str)
    liste = Path("data", "Liste.xlsx")
    output = Path("results", f"report_{month_str}.csv")
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        ok = process_month(month_dir, liste, output)
    except Exception as exc:
        log(f"Fehler: {exc}")
        raise
    else:
        if ok:
            log(f"Monat {month_str} erfolgreich verarbeitet.")
        else:
            log(f"Monat {month_str} mit Fehlern verarbeitet.")


if __name__ == "__main__":
    main()
