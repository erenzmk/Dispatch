from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from . import analyze_month, process_reports, summarize_by_id


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Dispatch helper CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    p_process = sub.add_parser("process", help="Process a day's reports")
    p_process.add_argument("day_dir", type=Path, help="Directory containing daily reports")
    p_process.add_argument("liste", type=Path, help="Path to Liste.xlsx")

    p_process_month = sub.add_parser("process-month", help="Process all day reports in a month")
    p_process_month.add_argument("month_dir", type=Path, help="Directory containing day folders")
    p_process_month.add_argument("liste", type=Path, help="Path to Liste.xlsx")

    p_analyze = sub.add_parser("analyze", help="Analyse a month of reports")
    p_analyze.add_argument("month_dir", type=Path, help="Directory containing day folders")
    p_analyze.add_argument("liste", type=Path, help="Path to Liste.xlsx")
    p_analyze.add_argument("-o", "--output", type=Path, default=Path("analysis.csv"), help="Output CSV file")

    p_all = sub.add_parser(
        "run-all",
        help="Process reports for a month and analyse them in one step",
    )
    p_all.add_argument("month_dir", type=Path, help="Directory containing day folders")
    p_all.add_argument("liste", type=Path, help="Path to Liste.xlsx")
    p_all.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("analysis.csv"),
        help="Output CSV file for analysis",
    )

    p_sum_id = sub.add_parser(
        "summarize-id",
        help="Summarize a report by technician ID",
    )
    p_sum_id.add_argument("excel_file", type=Path, help="Path to the Excel report")
    p_sum_id.add_argument("liste", type=Path, help="Path to Liste.xlsx")
    p_sum_id.add_argument("-o", "--output", type=Path, help="Output CSV file")

    args = parser.parse_args(argv)

    if args.command == "process":
        process_reports.main([str(args.day_dir), str(args.liste)])
    elif args.command == "process-month":
        process_reports.process_month(args.month_dir, args.liste)
    elif args.command == "analyze":
        analyze_month.main([str(args.month_dir), str(args.liste), "-o", str(args.output)])
    elif args.command == "run-all":
        process_reports.process_month(args.month_dir, args.liste)
        analyze_month.main([str(args.month_dir), str(args.liste), "-o", str(args.output)])
    elif args.command == "summarize-id":
        summary = summarize_by_id.summarize_report(args.excel_file, args.liste)
        df = pd.DataFrame(summary)
        if args.output:
            df.to_csv(args.output, index=False)
        else:
            print(df.to_string(index=False))


if __name__ == "__main__":
    main()
