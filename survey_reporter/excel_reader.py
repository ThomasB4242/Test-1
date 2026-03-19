"""Read an Excel toplines table into QuestionResult objects.

Expected format (one row per response option):
    Column A  variable      — SPSS-style variable name  (e.g. "issue_concern")
    Column B  question      — Question text / label
    Column C  response      — Response option label     (e.g. "Very concerned")
    Column D  n             — Count (integer)
    Column E  pct           — Percentage 0–100          (e.g. 30.5)

The first row must be a header row (column names are case-insensitive).
Rows for the same variable are collected into one QuestionResult, preserving
the order they appear in the file.

Alternatively the file may contain a sheet per section; the reader processes
all sheets by default (pass sheet_name to target one).

Usage
-----
    from survey_reporter.excel_reader import load_toplines
    results = load_toplines("toplines.xlsx")
    # then pass to build_pptx just like SPSS results
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
