"""Styled matplotlib chart renderers that return PNG BytesIO objects."""
from __future__ import annotations

import io
import textwrap
from typing import List

import matplotlib
matplotlib.use("Agg")
import matplotlib.lines as mlines
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt

from . import theme


# ---------------------------------------------------------------------------
# Render dimensions
# ---------------------------------------------------------------------------
CHART_FIG_W: float = 5.50
CHART_FIG_H: float = 5.20
CHART_DPI: int = 150

# Subplot margins (fractions of figure; matplotlib measures from bottom).
# Exported so slide_builder.py can map bar positions to slide coordinates.
AXIS_LEFT        = 0.30    # simple bar — short response labels
GRID_AXIS_LEFT   = 0.40    # grid / stacked — longer statement labels
AXIS_RIGHT       = 0.90
AXIS_TOP         = 0.95    # leave a sliver at top
AXIS_BOTTOM      = 0.03    # simple bar — no x-axis labels
GRID_AXIS_BOTTOM = 0.12    # grid — room for x-axis tick labels


def _hex_to_rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


def _text_color_for_bg(hex_bg: str) -> str:
    """White on dark backgrounds, gray on light ones."""
    return "white" if hex_bg in {theme.TEAL_DARK, theme.CORAL} else theme.GRAY_DARK


def _wrap_label(text: str, max_chars: int = 32) -> str:
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

    Labels displayed top-to-bottom in the order supplied.
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
    ax.set_xlim(0, max_val * 1.12 + 3)

    for i, (label, val, clr) in enumerate(zip(labels, values, bar_colors)):
        ax.barh(i, val, height=0.60, color=_hex_to_rgb(clr), linewidth=0)
        if val >= 8:
            ax.text(val / 2, i, f"{int(round(val))}",
                    ha="center", va="center",
                    fontsize=10, color="white", fontweight="bold")
        else:
            ax.text(val + 0.8, i, f"{int(round(val))}",
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
    axis_bottom: float | None = None,
) -> io.BytesIO:
    """Render a stacked 0–100 horizontal bar chart for grid questions.

    Parameters
    ----------
    row_labels   : One label per row, index 0 at the TOP.
    segments     : List of ``{label, values}`` dicts (one per scale option).
    colors       : Hex colour per segment.
    fig_h        : Figure height (inches). Defaults to CHART_FIG_H.
    axis_left    : Left subplot margin fraction. Defaults to GRID_AXIS_LEFT.
    axis_bottom  : Bottom subplot margin fraction. Defaults to GRID_AXIS_BOTTOM.
    """
    if colors is None:
        colors = theme.LIKERT_COLORS
    if fig_h is None:
        fig_h = CHART_FIG_H
    if axis_left is None:
        axis_left = GRID_AXIS_LEFT
    if axis_bottom is None:
        axis_bottom = GRID_AXIS_BOTTOM

    n_rows = len(row_labels)
    wrapped_labels = [_wrap_label(lbl) for lbl in row_labels]
    label_fs = 9 if n_rows <= 6 else 8

    fig, ax = plt.subplots(figsize=(CHART_FIG_W, fig_h))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.subplots_adjust(
        left=axis_left, right=AXIS_RIGHT,
        top=AXIS_TOP,   bottom=axis_bottom,
    )

    lefts = [0.0] * n_rows
    y_positions = list(range(n_rows))

    for i, seg in enumerate(segments):
        clr  = colors[i % len(colors)]
        vals = seg["values"]
        bars = ax.barh(
            y_positions, vals, left=lefts,
            height=0.55, color=_hex_to_rgb(clr), linewidth=0,
            zorder=3,
        )
        txt_clr = _text_color_for_bg(clr)
        for bar, val in zip(bars, vals):
            if val >= 7:
                cx = bar.get_x() + bar.get_width() / 2
                cy = bar.get_y() + bar.get_height() / 2
                ax.text(cx, cy, f"{int(round(val))}",
                        ha="center", va="center",
                        fontsize=9, color=txt_clr, fontweight="bold", zorder=4)
        lefts = [l + v for l, v in zip(lefts, vals)]

    # ── x-axis: grid lines at 0/25/50/75/100, labels, no tick marks ──────────
    ax.set_xlim(0, 100)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xticklabels(["0", "25", "50", "75", "100"])
    ax.tick_params(axis="x", length=0, labelsize=7,
                   labelcolor=theme.GRAY_DARK, pad=3)
    ax.xaxis.grid(True, linestyle="-", linewidth=0.6,
                  color="#E0E0E0", zorder=0)
    ax.set_axisbelow(True)

    # Bold bottom spine only
    for name, spine in ax.spines.items():
        if name == "bottom":
            spine.set_visible(True)
            spine.set_linewidth(1.4)
            spine.set_color(theme.GRAY_DARK)
        else:
            spine.set_visible(False)

    ax.set_yticks(y_positions)
    ax.set_yticklabels(wrapped_labels, fontsize=label_fs, color=theme.GRAY_DARK,
                       linespacing=1.1)
    ax.set_ylim(-0.55, n_rows - 0.45)
    ax.invert_yaxis()
    ax.tick_params(axis="y", length=0)  # no y tick marks either

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Legend strip
# ---------------------------------------------------------------------------

def legend_image(
    segment_labels: List[str],
    colors: List[str] | None = None,
    fig_w: float = CHART_FIG_W,
    circle_indices: set | None = None,
) -> io.BytesIO:
    """Render a small horizontal legend strip.

    Parameters
    ----------
    circle_indices : set of int
        Indices of entries that should use a circle marker instead of a
        colour rectangle (e.g. {0} for the 'Total agree' entry).
    """
    if colors is None:
        colors = theme.LIKERT_COLORS
    if circle_indices is None:
        circle_indices = set()

    n = len(segment_labels)
    handles = []
    for i, label in enumerate(segment_labels):
        clr = _hex_to_rgb(colors[i % len(colors)])
        if i in circle_indices:
            h = mlines.Line2D(
                [], [], color="none", marker="o",
                markerfacecolor=clr, markeredgecolor="none",
                markersize=9, label=label,
            )
        else:
            h = mpatches.Patch(color=clr, label=label)
        handles.append(h)

    fig, ax = plt.subplots(figsize=(fig_w, 0.30))
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")
    ax.set_axis_off()
    ax.legend(handles=handles, loc="center", ncol=n, fontsize=7,
              frameon=False, handlelength=1.2, handleheight=0.8,
              borderpad=0, columnspacing=0.9)
    plt.tight_layout(pad=0)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI,
                bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
