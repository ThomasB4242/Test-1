#!/usr/bin/env python3
"""
survey-report — automate reports from SPSS .sav files.

Usage:
    # Generate a PowerPoint (default):
    python main.py survey.sav
    python main.py survey.sav --output report.pptx --title "Energy Survey"

    # Scaffold an annotations YAML then use it:
    python main.py survey.sav --generate-annotations annotations.yaml
    python main.py survey.sav --annotations annotations.yaml --output report.pptx

    # Generate a PDF instead:
    python main.py survey.sav --format pdf
"""

import argparse
import sys
from pathlib import Path

from survey_reporter.reader import load_survey
from survey_reporter.analyzer import analyze


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="survey-report",
        description="Generate a survey report (PowerPoint or PDF) from an SPSS .sav file.",
    )
    parser.add_argument("sav_file", help="Path to the SPSS .sav file")
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output file path (default: <sav_file stem>_report.pptx or .pdf)",
    )
    parser.add_argument(
        "--format", "-f",
        choices=["pptx", "pdf"],
        default="pptx",
        help="Output format: pptx (default) or pdf",
    )
    parser.add_argument(
        "--title", "-t",
        default=None,
        help="Report title (default: derived from file name)",
    )
    parser.add_argument(
        "--annotations", "-a",
        default=None,
        metavar="YAML_FILE",
        help="Path to a YAML annotations file (headings, question text, bullet notes)",
    )
    parser.add_argument(
        "--generate-annotations",
        default=None,
        metavar="OUT_YAML",
        help="Scaffold an annotations YAML template from the .sav file and exit",
    )
    parser.add_argument(
        "--logo",
        default=None,
        metavar="IMAGE_FILE",
        help="Path to a logo image to embed in slides",
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

    # --generate-annotations: scaffold a YAML template and exit
    if args.generate_annotations:
        from survey_reporter.annotations import save_template
        save_template(results, args.generate_annotations, title=title)
        return 0

    # --- PDF path ---
    if args.format == "pdf":
        from survey_reporter.report import build_report
        output_path = Path(args.output) if args.output else sav_path.with_name(
            sav_path.stem + "_report.pdf"
        )
        print(f"Building PDF → {output_path}")
        build_report(results=results, output_path=output_path,
                     title=title, n_responses=meta.n_rows)
        print(f"Done!    : {output_path}")
        return 0

    # --- PowerPoint path (default) ---
    from survey_reporter.annotations import load_annotations, auto_annotations
    from survey_reporter.slide_builder import build_pptx

    output_path = Path(args.output) if args.output else sav_path.with_name(
        sav_path.stem + "_report.pptx"
    )

    if args.annotations:
        print(f"Annotations: {args.annotations}")
        ann = load_annotations(args.annotations)
    else:
        print("Annotations: auto-generated (no YAML supplied)")
        ann = auto_annotations(results, title=title)

    print(f"Building PPTX → {output_path}")
    build_pptx(
        results=results,
        annotations=ann,
        output_path=str(output_path),
        logo_path=args.logo,
    )

    print(f"Done!    : {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
