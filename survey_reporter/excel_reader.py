"""Read an Excel toplines table into QuestionResult objects.

Two formats are supported:

1. Flat format (load_toplines)
   Column A  variable  — SPSS-style variable name
   Column B  question  — Question text / label
   Column C  response  — Response option label
   Column D  n         — Count (integer)
   Column E  pct       — Percentage 0–100

2. Matrix / output toplines format (load_toplines_matrix)
   The formatted output produced by crosstab software (Q, SPSS, etc.).
   Each table starts with "Back to TOC" sentinel, followed by question text,
   column headers, data rows, and a "base n = N" footer.
   Two sub-types:
     - "grid_rows"  : each data row → one QuestionResult (grid battery)
     - "single"     : all data rows → one QuestionResult (single question)

Usage
-----
    from survey_reporter.excel_reader import load_toplines, load_toplines_matrix
    results = load_toplines("flat_toplines.xlsx")
    results = load_toplines_matrix("output_toplines.xlsx", VARIABLE_MAP)
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from .analyzer import QuestionResult, FrequencyRow


def load_toplines(
    path: str | Path,
    sheet_name: str | int | None = None,
    total_n: Optional[int] = None,
) -> list[QuestionResult]:
    """Parse an Excel toplines file and return a list of QuestionResult objects.

    Parameters
    ----------
    path        : Path to the .xlsx file.
    sheet_name  : Sheet name or index to read.  ``None`` reads all sheets
                  (rows concatenated in sheet order).
    total_n     : Override total base n for percentage calculations.
                  If not supplied, it is inferred from the largest group sum.

    Returns
    -------
    list[QuestionResult]  — one entry per unique variable in the file.
    """
    try:
        import openpyxl
    except ImportError:
        raise ImportError(
            "openpyxl is required to read Excel files.  "
            "Install it with:  pip install openpyxl"
        )

    wb = openpyxl.load_workbook(path, data_only=True)

    if sheet_name is not None:
        sheets = [wb[sheet_name] if isinstance(sheet_name, str)
                  else wb.worksheets[sheet_name]]
    else:
        sheets = wb.worksheets

    # Collect raw rows across all target sheets
    raw_rows: list[dict] = []
    for ws in sheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # First row is the header
        header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
        col = _map_columns(header)
        for row in rows[1:]:
            if all(c is None for c in row):
                continue  # skip blank rows
            raw_rows.append({
                "variable": _cell(row, col.get("variable")),
                "question": _cell(row, col.get("question")),
                "response": _cell(row, col.get("response")),
                "n":        _num(row, col.get("n")),
                "pct":      _num(row, col.get("pct")),
            })

    # Group by variable (preserving order)
    from collections import OrderedDict
    groups: OrderedDict[str, list[dict]] = OrderedDict()
    for r in raw_rows:
        var = str(r["variable"]).strip() if r["variable"] else ""
        if not var:
            continue
        groups.setdefault(var, []).append(r)

    results: list[QuestionResult] = []
    for var, rows in groups.items():
        question_label = rows[0]["question"] or var
        frequencies    = []
        for row in rows:
            resp = str(row["response"]).strip() if row["response"] else ""
            pct  = float(row["pct"]) if row["pct"] is not None else 0.0
            n    = int(row["n"])    if row["n"]   is not None else 0
            if resp:
                frequencies.append(FrequencyRow(label=resp, code=resp, count=n, percent=pct))
        if frequencies:
            results.append(QuestionResult(
                variable=var,
                label=str(question_label),
                type="categorical",
                frequencies=frequencies,
                n_valid=sum(f.count for f in frequencies),
            ))

    return results


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_COL_ALIASES = {
    "variable":  {"variable", "var", "varname", "variable_name", "q", "qcode"},
    "question":  {"question", "question_text", "label", "qlabel", "q_label",
                  "question_label"},
    "response":  {"response", "response_label", "option", "answer", "value",
                  "response_option"},
    "n":         {"n", "count", "freq", "frequency"},
    "pct":       {"pct", "percent", "percentage", "%", "pct_col"},
}


def _map_columns(header: list[str]) -> dict[str, int]:
    """Map canonical column names to positional indices."""
    result: dict[str, int] = {}
    for i, col in enumerate(header):
        for canonical, aliases in _COL_ALIASES.items():
            if col in aliases and canonical not in result:
                result[canonical] = i
                break
    # Fallback: if columns not found by name, try positional (A=var, B=q, C=resp, D=n, E=pct)
    for pos, canonical in enumerate(["variable", "question", "response", "n", "pct"]):
        if canonical not in result and pos < len(header):
            result[canonical] = pos
    return result


def _cell(row: tuple, idx: int | None):
    if idx is None or idx >= len(row):
        return None
    v = row[idx]
    return str(v).strip() if v is not None else None


def _num(row: tuple, idx: int | None) -> float | None:
    if idx is None or idx >= len(row):
        return None
    v = row[idx]
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Matrix / output toplines parser
# ---------------------------------------------------------------------------

def load_toplines_matrix(
    path: str | Path,
    variable_map: dict,
    sheet_name: str = "Tables",
) -> list[QuestionResult]:
    """Parse a formatted output-toplines Excel file into QuestionResult objects.

    Parameters
    ----------
    path         : Path to the .xlsx file.
    variable_map : Dict mapping a partial question-text string (lower-case match)
                   to a table spec.  Two spec types are supported:

                   ``type="grid_rows"`` — each data row → one QuestionResult::

                       {
                           "type": "grid_rows",
                           "row_vars": {
                               "Coal": "energy_climate_coal",
                               ...
                           },
                           "label_map": {"Coal": "Coal"},   # optional display labels
                           "question_label": "...",
                           "n": 1011,
                       }

                   ``type="single"`` — all data rows → one QuestionResult::

                       {
                           "type": "single",
                           "variable": "energy_concern",
                           "question_label": "...",
                           "n": 1011,
                           "exclude_rows": ["TOTAL SUPPORT", ...],  # optional
                           "label_map": {"long label...": "Short label"},  # optional
                       }

    sheet_name   : Sheet in the workbook that contains the tables (default "Tables").

    Notes
    -----
    Table-boundary detection: tables are separated by rows whose first cell equals
    "Back to TOC" (case-insensitive).  Each table then has:
      Row +1 : question text
      Row +2 : column headers ("Row %" + scale options, or " " + "%")
      Rows +3…: data rows (label + values)
      Final  : "base n = N" row

    For ``type="single"`` the frequencies are stored in *reverse file order* so
    that the existing slide-builder (which calls ``list(reversed(freqs))``) will
    display the first file row at the top of the chart.
    """
    try:
        import openpyxl
    except ImportError:
        raise ImportError("openpyxl is required.  pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    all_rows = list(ws.iter_rows(values_only=True))

    # Locate the start of each table ("Back to TOC" sentinel)
    table_starts = [
        i for i, row in enumerate(all_rows)
        if row[0] is not None and str(row[0]).strip().lower() == "back to toc"
    ]

    results: list[QuestionResult] = []

    for t_start in table_starts:
        # Row +1 : question text (may span merged cells — col 0 carries it)
        q_text_row = all_rows[t_start + 1] if t_start + 1 < len(all_rows) else (None,)
        q_text = str(q_text_row[0]).strip() if q_text_row[0] is not None else ""

        # Match against variable_map keys
        spec = None
        for key, s in variable_map.items():
            if key.lower() in q_text.lower():
                spec = s
                break
        if spec is None:
            continue

        # Row +2 : column headers
        col_row = all_rows[t_start + 2] if t_start + 2 < len(all_rows) else ()
        col_headers = [
            str(c).strip() if c is not None else ""
            for c in col_row
        ]
        is_wide = col_headers[0].lower() == "row %"

        # Data rows: t_start+3 onward until "base n = ..." or blank section
        exclude_set = {s.lower() for s in spec.get("exclude_rows", [])}
        label_map   = spec.get("label_map", {})
        n           = int(spec.get("n", 1011))
        q_label     = spec.get("question_label", q_text)

        data_rows: list[tuple] = []
        for row in all_rows[t_start + 3:]:
            if all(c is None for c in row):
                break
            cell0 = str(row[0]).strip() if row[0] is not None else ""
            if cell0.lower().startswith("base n"):
                break
            if not cell0:
                break
            data_rows.append(row)

        if spec["type"] == "grid_rows":
            row_vars = spec["row_vars"]   # {file_row_label: variable_name}
            for dr in data_rows:
                row_label = str(dr[0]).strip() if dr[0] is not None else ""
                if row_label.lower() in exclude_set:
                    continue
                var_name = row_vars.get(row_label)
                if var_name is None:
                    continue
                display_label = label_map.get(row_label, row_label)
                freqs: list[FrequencyRow] = []
                # Columns: col_headers[1:] = scale options
                for i, scale_opt in enumerate(col_headers[1:], start=1):
                    if not scale_opt or i >= len(dr):
                        continue
                    pct = float(dr[i]) if dr[i] is not None else 0.0
                    cnt = int(round(pct * n / 100))
                    freqs.append(FrequencyRow(
                        label=scale_opt, code=scale_opt,
                        count=cnt, percent=round(pct, 1),
                    ))
                if freqs:
                    results.append(QuestionResult(
                        variable=var_name,
                        label=display_label,
                        type="categorical",
                        frequencies=freqs,
                        n_valid=n,
                    ))

        elif spec["type"] == "single":
            var_name = spec["variable"]
            freqs = []
            for dr in data_rows:
                row_label = str(dr[0]).strip() if dr[0] is not None else ""
                if not row_label:
                    continue
                if row_label.lower() in exclude_set:
                    continue
                display_label = label_map.get(row_label, row_label)
                pct = float(dr[1]) if len(dr) > 1 and dr[1] is not None else 0.0
                cnt = int(round(pct * n / 100))
                freqs.append(FrequencyRow(
                    label=display_label, code=display_label,
                    count=cnt, percent=round(pct, 1),
                ))
            # Reverse so that list(reversed(freqs)) in slide_builder puts
            # the first file row at the TOP of the chart.
            freqs = list(reversed(freqs))
            if freqs:
                results.append(QuestionResult(
                    variable=var_name,
                    label=q_label,
                    type="categorical",
                    frequencies=freqs,
                    n_valid=n,
                ))

    return results
