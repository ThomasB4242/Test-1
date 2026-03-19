"""Styled matplotlib chart renderers that return PNG BytesIO objects."""
from __future__ import annotations

import io
import textwrap
from typing import List

import matplotlib
matplotlib.use("Agg")
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt

from . import theme


# ---------------------------------------------------------------------------
# Render dimensions — must match theme.CHART_BOX width so the image is
# embedded without horizontal distortion.  Height is passed dynamically.
# ---------------------------------------------------------------------------
CHART_FIG_W: float = 5.50   # inches  (= theme.CHART_BOX[2])
CHART_FIG_H: float = 5.20   # default figure height when not overridden
CHART_DPI: int = 150

# Explicit subplot margins (fractions of figure; matplotlib measures from bottom).
# Exported so slide_builder.py can compute bar y-positions in slide coordinates.
AXIS_LEFT        = 0.30   # simple bar — short response labels (e.g. "Excellent")
GRID_AXIS_LEFT   = 0.40   # grid / stacked — longer statement labels
AXIS_RIGHT       = 0.90   # fraction from left
AXIS_TOP         = 0.97   # fraction from bottom
AXIS_BOTTOM      = 0.03   # fraction from bottom


def _hex_to_rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


def _text_color_for_bg(hex_bg: str) -> str:
    """Return 'white' for dark backgrounds, GRAY_DARK for light ones."""
    dark_bgs = {theme.TEAL_DARK, theme.CORAL}
    return "white" if hex_bg in dark_bgs else theme.GRAY_DARK


def _wrap_label(text: str, max_chars: int = 32) -> str:
    """Wrap text at word boundaries for use as a y-axis tick label."""
    return "\n".join(textwrap.wrap(text, max_chars))


# ---------------------------------------------------------------------------
# Simple horizontal bar chart (one bar per response option)
# ---------------------------------------------------------------------------

def render_simple_bar(
    labels: List[str],
    values: List[float],
    color: str | None = None,
    bar_colors: List[str] | None = None,
    fig_h: float | None = None,
) -> io.BytesIO:
    """Render a simple horizontal bar chart.

    Labels are displayed top-to-bottom in the order supplied.
    Pass them in the desired display order (e.g. most-positive first).
    """
    if bar_colors is None:
        bar_colors = [color or theme.BAR_COLOR] * len(labels)
    if fig_h is None:
        fig_h = CHART_FIG_H

    n = len(labels)
    fig, ax = plt.subplots(figsize=(CHART_FIG_W, fig_h))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.subplots_adjust(
        left=AXIS_LEFT, right=AXIS_RIGHT,
        top=AXIS_TOP,   bottom=AXIS_BOTTOM,
    )

    max_val = max(values) if values else 1
    ax.set_xlim(0, max_val * 1.10 + 3)

    bar_h = 0.60
    for i, (label, val, clr) in enumerate(zip(labels, values, bar_colors)):
        ax.barh(i, val, height=bar_h, color=_hex_to_rgb(clr), linewidth=0)
        if val >= 8:
            ax.text(val / 2, i, f"{int(round(val))}",
                    ha="center", va="center",
                    fontsize=10, color="white", fontweight="bold")
        else:
            ax.text(val + 1.0, i, f"{int(round(val))}",
                    ha="left", va="center",
                    fontsize=10, color=theme.GRAY_DARK, fontweight="bold")

    ax.set_yticks(list(range(n)))
    ax.set_yticklabels(labels, fontsize=10, color=theme.GRAY_DARK)
    ax.set_ylim(-0.55, n - 0.45)
    ax.invert_yaxis()
    ax.xaxis.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    buf = io.BytesIO()
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
    fig_h: float | None = None,
    axis_left: float | None = None,
) -> io.BytesIO:
    """Render a stacked horizontal bar chart for grid / multi-row questions.

    Parameters
    ----------
    row_labels : list[str]
        One label per row, index 0 at the TOP.
    segments : list[dict]
        Each dict has ``label`` (str) and ``values`` (list[float], 0–100).
    colors : list[str] | None
        Hex colour per segment.
    fig_h : float | None
        Figure height in inches. Defaults to CHART_FIG_H.
    axis_left : float | None
        Left subplot margin (fraction).  Use GRID_AXIS_LEFT for long labels.
    """
    if colors is None:
        colors = theme.LIKERT_COLORS
    if fig_h is None:
        fig_h = CHART_FIG_H
    if axis_left is None:
        axis_left = GRID_AXIS_LEFT

    n_rows = len(row_labels)
    # Wrap long labels for legibility
    wrapped_labels = [_wrap_label(lbl) for lbl in row_labels]

    # Scale font size with number of rows
    label_fs = 9 if n_rows <= 6 else 8

    fig, ax = plt.subplots(figsize=(CHART_FIG_W, fig_h))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.subplots_adjust(
        left=axis_left, right=AXIS_RIGHT,
        top=AXIS_TOP,   bottom=AXIS_BOTTOM,
    )

    lefts = [0.0] * n_rows
    y_positions = list(range(n_rows))
    bar_h = 0.55

    for i, seg in enumerate(segments):
        clr = colors[i % len(colors)]
        vals = seg["values"]
        bars = ax.barh(
            y_positions, vals, left=lefts,
            height=bar_h, color=_hex_to_rgb(clr), linewidth=0,
        )
        txt_clr = _text_color_for_bg(clr)
        for bar, val in zip(bars, vals):
            if val >= 6:
                cx = bar.get_x() + bar.get_width() / 2
                cy = bar.get_y() + bar.get_height() / 2
                ax.text(cx, cy, f"{int(round(val))}",
                        ha="center", va="center",
                        fontsize=9, color=txt_clr, fontweight="bold")
        lefts = [l + v for l, v in zip(lefts, vals)]

    ax.set_yticks(y_positions)
    ax.set_yticklabels(wrapped_labels, fontsize=label_fs, color=theme.GRAY_DARK,
                       linespacing=1.1)
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.55, n_rows - 0.45)
    ax.invert_yaxis()
    ax.xaxis.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI)
    plt.close(fig)
    buf.seek(0)
    return buf


def legend_image(segment_labels: List[str], colors: List[str] | None = None,
                 fig_w: float = CHART_FIG_W) -> io.BytesIO:
    """Render a small legend strip."""
    if colors is None:
        colors = theme.LIKERT_COLORS
    n = len(segment_labels)
    fig, ax = plt.subplots(figsize=(fig_w, 0.30))
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
              borderpad=0, columnspacing=0.8)
    plt.tight_layout(pad=0)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
