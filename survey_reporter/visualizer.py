"""Generate chart images for survey questions."""

from __future__ import annotations

import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

from .analyzer import QuestionResult

# ── Styling ───────────────────────────────────────────────────────────────────
PALETTE = sns.color_palette("Blues_r", 10)
sns.set_theme(style="whitegrid", palette=PALETTE, font_scale=1.0)

BAR_COLOR = "#2E75B6"
HIST_COLOR = "#2E75B6"
FIGSIZE_CAT = (7, None)   # width fixed; height computed from n options
FIGSIZE_NUM = (6, 3.5)

_tmpdir = Path(tempfile.mkdtemp(prefix="survey_charts_"))


def chart_for_question(result: QuestionResult) -> Path | None:
    """Render the appropriate chart and return the path to a PNG file.

    Returns None if there is nothing meaningful to chart (e.g. < 2 values).
    """
    if result.type == "categorical":
        return _bar_chart(result)
    else:
        return _numeric_chart(result)


def _bar_chart(result: QuestionResult) -> Path | None:
    freqs = result.frequencies
    if not freqs:
        return None

    labels = [r.label for r in freqs]
    percents = [r.percent for r in freqs]

    # Dynamic height: ~0.45 inch per bar, min 2 inches
    height = max(2.0, 0.45 * len(labels) + 0.8)
    fig, ax = plt.subplots(figsize=(FIGSIZE_CAT[0], height))

    bars = ax.barh(labels, percents, color=BAR_COLOR, edgecolor="white", height=0.6)

    # Annotate bars with percentage labels
    for bar, pct in zip(bars, percents):
        ax.text(
            bar.get_width() + 0.5,
            bar.get_y() + bar.get_height() / 2,
            f"{pct:.1f}%",
            va="center",
            fontsize=8,
            color="#333333",
        )

    ax.set_xlabel("Percentage (%)", fontsize=9)
    ax.set_xlim(0, min(100, max(percents) * 1.2 + 5))
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%g%%"))
    ax.invert_yaxis()
    ax.tick_params(axis="y", labelsize=8)
    ax.spines[["top", "right"]].set_visible(False)

    plt.tight_layout()
    out = _tmpdir / f"{result.variable}_bar.png"
    fig.savefig(out, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return out


def _numeric_chart(result: QuestionResult) -> Path | None:
    # We don't have the raw series here, so we draw a summary stats figure
    stats = {
        "Mean": result.mean,
        "Median": result.median,
        "Std Dev": result.std,
        "Min": result.min,
        "Max": result.max,
    }
    if any(v is None for v in stats.values()):
        return None

    fig, ax = plt.subplots(figsize=FIGSIZE_NUM)
    keys = list(stats.keys())
    vals = [stats[k] for k in keys]

    bars = ax.bar(keys, vals, color=BAR_COLOR, edgecolor="white", width=0.5)
    for bar, val in zip(bars, vals):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.01 * (max(vals) or 1),
            f"{val:.2f}",
            ha="center",
            va="bottom",
            fontsize=8,
            color="#333333",
        )

    ax.set_ylabel("Value", fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="x", labelsize=9)
    plt.tight_layout()

    out = _tmpdir / f"{result.variable}_stats.png"
    fig.savefig(out, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return out
