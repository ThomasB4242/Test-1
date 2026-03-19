#!/usr/bin/env python3
"""
survey-report — automate PDF reports from SPSS .sav files.

Usage:
    python main.py survey.sav
    python main.py survey.sav --output report.pdf --title "Customer Satisfaction Q1 2026"
"""

import argparse
import sys
from pathlib import Path

from survey_reporter.reader import load_survey
from survey_reporter.analyzer import analyze
from survey_reporter.report import build_report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="survey-report",
        description="Generate a PDF survey report from an SPSS .sav file.",
    )
    parser.add_argument("sav_file", help="Path to the SPSS .sav file")
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output PDF path (default: <sav_file stem>_report.pdf)",
    )
    parser.add_argument(
        "--title", "-t",
        default=None,
        help='Report title (default: derived from file name)',
    )
    parser.add_argument(
        "--skip", "-s",
        nargs="*",
        default=[],
        metavar="VAR",
        help="Variable names to exclude from the report",
    )

    args = parser.parse_args(argv)

    sav_path = Path(args.sav_file)

    # Defaults
    output_path = Path(args.output) if args.output else sav_path.with_name(
        sav_path.stem + "_report.pdf"
    )
    title = args.title or sav_path.stem.replace("_", " ").replace("-", " ").title()

    print(f"Loading  : {sav_path}")
    try:
        df, meta = load_survey(sav_path)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Rows     : {meta.n_rows:,}")
    print(f"Variables: {len(meta.column_names)}")

    # Filter skipped variables
    skip_set = set(args.skip or [])
    if skip_set:
        keep = [c for c in meta.column_names if c not in skip_set]
        df = df[keep]
        meta.column_names = keep
        print(f"Skipping : {', '.join(skip_set)}")

    print("Analysing questions …")
    results = analyze(df, meta)

    print(f"Building PDF → {output_path}")
    build_report(
        results=results,
        output_path=output_path,
        title=title,
        n_responses=meta.n_rows,
    )

    print(f"Done!    : {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
