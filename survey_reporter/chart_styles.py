"""Styled matplotlib chart renderers that return PNG BytesIO objects."""
from __future__ import annotations

import io
from typing import List

import matplotlib
matplotlib.use("Agg")
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt

from . import theme


# ---------------------------------------------------------------------------
# Render dimensions — must match theme.CHART_BOX width/height so the image
# is embedded at exactly the right size without distortion.
# ---------------------------------------------------------------------------
CHART_FIG_W: float = 5.50   # inches  (= theme.CHART_BOX[2])
CHART_FIG_H: float = 5.30   # inches  (= theme.CHART_BOX[3])
CHART_DPI: int = 150

# Explicit subplot margins (fractions of figure; matplotlib measures from bottom).
# Exported so slide_builder.py can compute bar y-positions in slide coordinates.
AXIS_LEFT   = 0.30   # fraction from left  (room for y-axis labels)
AXIS_RIGHT  = 0.90   # fraction from left  (right margin)
AXIS_TOP    = 0.97   # fraction from bottom (axis top)
AXIS_BOTTOM = 0.03   # fraction from bottom (axis bottom)


def _hex_to_rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# Simple horizontal bar chart (one bar per response option)
# ---------------------------------------------------------------------------

def render_simple_bar(
    labels: List[str],
    values: List[float],
    color: str | None = None,
) -> io.BytesIO:
    """Render a simple horizontal bar chart.

    Labels are displayed top-to-bottom in the order supplied.
    Pass them in the desired display order (e.g. most-positive first).

    Parameters
    ----------
    labels : list[str]
        Bar labels, index 0 at the TOP of the chart.
    values : list[float]
        Percentages 0–100, matching *labels*.
    color : str | None
        Hex colour.  Defaults to BAR_COLOR.

    Returns
    -------
    io.BytesIO  PNG image data.
    """
    if color is None:
        color = theme.BAR_COLOR

    n = len(labels)
    fig, ax = plt.subplots(figsize=(CHART_FIG_W, CHART_FIG_H))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.subplots_adjust(
        left=AXIS_LEFT, right=AXIS_RIGHT,
        top=AXIS_TOP, bottom=AXIS_BOTTOM,
    )

    y_positions = list(range(n))
    bars = ax.barh(
        y_positions,
        values,
        height=0.65,
        color=_hex_to_rgb(color),
        linewidth=0,
    )

    max_val = max(values) if values else 1
    ax.set_xlim(0, max_val * 1.08 + 4)

    for bar, val in zip(bars, values):
        w = bar.get_width()
        cy = bar.get_y() + bar.get_height() / 2
        if val >= 8:
            ax.text(w / 2, cy, f"{int(round(val))}",
                    ha="center", va="center",
                    fontsize=10, color="white", fontweight="bold")
        else:
            ax.text(w + 1.0, cy, f"{int(round(val))}",
                    ha="left", va="center",
                    fontsize=10, color=theme.GRAY_DARK, fontweight="bold")

    ax.set_yticks(y_positions)
    ax.set_yticklabels(labels, fontsize=10, color=theme.GRAY_DARK)
    ax.set_ylim(-0.55, n - 0.45)
    ax.invert_yaxis()   # index 0 displayed at TOP
    ax.xaxis.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    buf = io.BytesIO()
    # Do NOT use bbox_inches='tight' — we want the figure at exactly CHART_FIG_W×H
    plt.savefig(buf, format="png", dpi=CHART_DPI)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Stacked horizontal bar chart (one row per grid item / survey statement)
# ---------------------------------------------------------------------------

def render_stacked_bar(
    row_labels: List[str],
    segments: List[dict],
    colors: List[str] | None = None,
) -> io.BytesIO:
    """Render a stacked horizontal bar chart for grid / multi-row questions.

    Parameters
    ----------
    row_labels : list[str]
        One label per ROW (survey item / statement), index 0 at the TOP.
    segments : list[dict]
        Ordered list of response-option dicts, each with:
          ``label`` (str) and ``values`` (list[float]) — one value per row,
          in the same order as *row_labels*.  Values are percentages 0–100.
    colors : list[str] | None
        Hex colours, one per segment.  Defaults to LIKERT_COLORS.

    Returns
    -------
    io.BytesIO  PNG image data.
    """
    if colors is None:
        colors = theme.LIKERT_COLORS

    n_rows = len(row_labels)

    fig, ax = plt.subplots(figsize=(CHART_FIG_W, CHART_FIG_H))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.subplots_adjust(
        left=AXIS_LEFT, right=AXIS_RIGHT,
        top=AXIS_TOP, bottom=AXIS_BOTTOM,
    )

    lefts = [0.0] * n_rows
    y_positions = list(range(n_rows))

    for i, seg in enumerate(segments):
        color = colors[i % len(colors)]
        vals = seg["values"]
        bars = ax.barh(
            y_positions,
            vals,
            left=lefts,
            height=0.65,
            color=_hex_to_rgb(color),
            linewidth=0,
        )
        for bar, val in zip(bars, vals):
            if val >= 6:
                cx = bar.get_x() + bar.get_width() / 2
                cy = bar.get_y() + bar.get_height() / 2
                txt_color = "white" if i < 2 or i == 3 else theme.GRAY_DARK
                ax.text(cx, cy, f"{int(round(val))}",
                        ha="center", va="center",
                        fontsize=9, color=txt_color, fontweight="bold")
        lefts = [l + v for l, v in zip(lefts, vals)]

    ax.set_yticks(y_positions)
    ax.set_yticklabels(row_labels, fontsize=10, color=theme.GRAY_DARK)
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.55, n_rows - 0.45)
    ax.invert_yaxis()   # index 0 displayed at TOP
    ax.xaxis.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI)
    plt.close(fig)
    buf.seek(0)
    return buf


def legend_image(segment_labels: List[str], colors: List[str] | None = None) -> io.BytesIO:
    """Render a small legend strip."""
    if colors is None:
        colors = theme.LIKERT_COLORS
    n = len(segment_labels)
    fig, ax = plt.subplots(figsize=(CHART_FIG_W, 0.3))
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")
    ax.set_axis_off()
    patches = [
        mpatches.Patch(color=_hex_to_rgb(colors[i % len(colors)]),
                       label=segment_labels[i])
        for i in range(n)
    ]
    ax.legend(handles=patches, loc="center", ncol=n, fontsize=7,
              frameon=False, handlelength=1.2, handleheight=0.8,
              borderpad=0, columnspacing=1.0)
    plt.tight_layout(pad=0)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
