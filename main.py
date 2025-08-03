from __future__ import annotations

import argparse
from pathlib import Path

from dispatch import analyze_month, aggregate_warnings, process_reports


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Dispatch helper CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    p_process = sub.add_parser("process", help="Process a day's reports")
    p_process.add_argument("day_dir", type=Path, help="Directory containing daily reports")
    p_process.add_argument("liste", type=Path, help="Path to Liste.xlsx")

    p_analyze = sub.add_parser("analyze", help="Analyse a month of reports")
    p_analyze.add_argument("month_dir", type=Path, help="Directory containing day folders")
    p_analyze.add_argument("liste", type=Path, help="Path to Liste.xlsx")
    p_analyze.add_argument("-o", "--output", type=Path, default=Path("analysis.csv"), help="Output CSV file")

    p_warnings = sub.add_parser("warnings", help="Aggregate unknown technician warnings")
    p_warnings.add_argument("report_dir", type=Path, help="Path to report directory")
    p_warnings.add_argument("--liste", type=Path, default=Path("Liste.xlsx"), help="Path to Liste.xlsx")

    args = parser.parse_args(argv)

    if args.command == "process":
        process_reports.main([str(args.day_dir), str(args.liste)])
    elif args.command == "analyze":
        analyze_month.main([str(args.month_dir), str(args.liste), "-o", str(args.output)])
    elif args.command == "warnings":
        aggregate_warnings.main([str(args.report_dir), "--liste", str(args.liste)])


if __name__ == "__main__":
    main()
