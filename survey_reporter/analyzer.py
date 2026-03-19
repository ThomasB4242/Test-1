"""Analyze survey questions and produce per-question statistics."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

import pandas as pd

from .reader import SurveyMeta


QuestionType = Literal["categorical", "numeric"]


@dataclass
class FrequencyRow:
    label: str
    code: Any
    count: int
    percent: float


@dataclass
class QuestionResult:
    variable: str
    label: str
    type: QuestionType
    n_valid: int          # non-null responses
    # Categorical fields
    frequencies: list[FrequencyRow] = field(default_factory=list)
    # Numeric fields
    mean: float | None = None
    median: float | None = None
    std: float | None = None
    min: float | None = None
    max: float | None = None


def analyze(df: pd.DataFrame, meta: SurveyMeta) -> list[QuestionResult]:
    """Return a QuestionResult for every column in *df*.

    Detection heuristic:
    - If the variable has SPSS value labels → categorical
    - If the series is numeric with no value labels → numeric
    - String columns with no value labels → categorical (use raw values)
    """
    results: list[QuestionResult] = []

    for var in meta.column_names:
        series = df[var].dropna()
        label = meta.variable_labels.get(var, var)
        n_valid = int(series.notna().sum()) if var in df.columns else 0
        series = df[var].dropna()

        if var in meta.value_labels and meta.value_labels[var]:
            result = _analyze_categorical(var, label, series, meta.value_labels[var])
        elif pd.api.types.is_numeric_dtype(series):
            result = _analyze_numeric(var, label, series)
        else:
            # String / open-ended — treat as categorical without code mapping
            result = _analyze_categorical(var, label, series, {})

        result.n_valid = len(series)
        results.append(result)

    return results


def _analyze_categorical(
    var: str,
    label: str,
    series: pd.Series,
    value_labels: dict[Any, str],
) -> QuestionResult:
    counts = series.value_counts(sort=False)
    total = counts.sum()

    frequencies: list[FrequencyRow] = []
    for code, count in counts.items():
        lbl = value_labels.get(code, str(code))
        frequencies.append(
            FrequencyRow(
                label=lbl,
                code=code,
                count=int(count),
                percent=round(100 * count / total, 1) if total else 0.0,
            )
        )

    # Sort by code if codes are numeric/ordinal, otherwise by frequency desc
    if value_labels:
        try:
            frequencies.sort(key=lambda r: float(r.code))
        except (TypeError, ValueError):
            pass
    else:
        frequencies.sort(key=lambda r: r.count, reverse=True)

    return QuestionResult(
        variable=var,
        label=label,
        type="categorical",
        n_valid=len(series),
        frequencies=frequencies,
    )


def _analyze_numeric(var: str, label: str, series: pd.Series) -> QuestionResult:
    return QuestionResult(
        variable=var,
        label=label,
        type="numeric",
        n_valid=len(series),
        mean=round(float(series.mean()), 2),
        median=round(float(series.median()), 2),
        std=round(float(series.std()), 2),
        min=round(float(series.min()), 2),
        max=round(float(series.max()), 2),
    )
