"""Load SPSS .sav files into a pandas DataFrame with metadata."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
import pyreadstat


@dataclass
class SurveyMeta:
    """Cleaned metadata extracted from a pyreadstat metadata object."""
    variable_labels: dict[str, str]          # varname -> human label
    value_labels: dict[str, dict[Any, str]]  # varname -> {code: label}
    column_names: list[str]
    n_rows: int


def load_survey(path: str | Path) -> tuple[pd.DataFrame, SurveyMeta]:
    """Read an SPSS .sav file and return (DataFrame, SurveyMeta).

    Parameters
    ----------
    path:
        Path to the .sav file.

    Returns
    -------
    df:
        One row per respondent; column names are SPSS variable names.
    meta:
        Human-readable labels and value mappings from the SPSS file.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"SPSS file not found: {path}")

    df, raw_meta = pyreadstat.read_sav(str(path), apply_value_formats=False)

    variable_labels: dict[str, str] = {
        k: v for k, v in raw_meta.column_names_to_labels.items() if v
    }
    # Fall back to variable name if no label exists
    for col in df.columns:
        variable_labels.setdefault(col, col)

    meta = SurveyMeta(
        variable_labels=variable_labels,
        value_labels=raw_meta.variable_value_labels,
        column_names=list(df.columns),
        n_rows=len(df),
    )

    return df, meta
