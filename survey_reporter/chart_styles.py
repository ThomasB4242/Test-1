"""Styled matplotlib chart renderers that return PNG BytesIO objects."""
from __future__ import annotations

import io
from typing import List

import matplotlib
matplotlib.use("Agg")
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from . import theme


def _hex_to_rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


def render_stacked_bar(
    categories: List[str],
    segments: List[dict],
    colors: List[str] | None = None,
) -> io.BytesIO:
    """Render a horizontal stacked bar chart.

    Parameters
    ----------
    categories : list[str]
        Row labels (one bar per category), bottom to top.
    segments : list[dict]
        Each dict has ``label`` (str) and ``values`` (list[float]) — one value per
        category, in the same order as *categories*.  Values are percentages (0–100).
    colors : list[str] | None
        Hex colours, one per segment.  Defaults to LIKERT_COLORS.

    Returns
    -------
    io.BytesIO  PNG image data.
    """
    if colors is None:
        colors = theme.LIKERT_COLORS

    n_cats = len(categories)
    bar_h = 0.55
    fig_h = max(1.5, n_cats * bar_h + 0.8)
    fig, ax = plt.subplots(figsize=(5.5, fig_h))
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")

    lefts = [0.0] * n_cats
    y_positions = list(range(n_cats))

    for i, seg in enumerate(segments):
        color = colors[i % len(colors)]
        vals = seg["values"]
        bars = ax.barh(
            y_positions,
            vals,
            left=lefts,
            height=bar_h * 0.9,
            color=_hex_to_rgb(color),
            linewidth=0,
        )
        for bar, val in zip(bars, vals):
            if val >= 6:
                cx = bar.get_x() + bar.get_width() / 2
                cy = bar.get_y() + bar.get_height() / 2
                # Choose text colour for contrast
                txt_color = "white" if i < 2 or i == 3 else theme.GRAY_DARK
                ax.text(
                    cx, cy,
                    f"{int(round(val))}",
                    ha="center", va="center",
                    fontsize=9, color=txt_color,
                    fontweight="bold",
                )
        lefts = [l + v for l, v in zip(lefts, vals)]

    ax.set_yticks(y_positions)
    ax.set_yticklabels(categories, fontsize=10, color=theme.GRAY_DARK)
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.5, n_cats - 0.5)
    ax.xaxis.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    plt.tight_layout(pad=0.2)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def render_simple_bar(
    labels: List[str],
    values: List[float],
    color: str | None = None,
) -> io.BytesIO:
    """Render a simple horizontal bar chart (one colour, % labels at right).

    Parameters
    ----------
    labels : list[str]
        Bar labels, bottom to top.
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
    bar_h = 0.55
    fig_h = max(1.5, n * bar_h + 0.8)
    fig, ax = plt.subplots(figsize=(5.5, fig_h))
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")

    y_positions = list(range(n))
    bars = ax.barh(
        y_positions,
        values,
        height=bar_h * 0.85,
        color=_hex_to_rgb(color),
        linewidth=0,
    )

    for bar, val in zip(bars, values):
        right = bar.get_width()
        cy = bar.get_y() + bar.get_height() / 2
        ax.text(
            right + 1.0, cy,
            f"{int(round(val))}",
            ha="left", va="center",
            fontsize=10, color=theme.GRAY_DARK,
            fontweight="bold",
        )

    ax.set_yticks(y_positions)
    ax.set_yticklabels(labels, fontsize=10, color=theme.GRAY_DARK)
    ax.set_xlim(0, max(values) * 1.25 + 5)
    ax.set_ylim(-0.5, n - 0.5)
    ax.xaxis.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    plt.tight_layout(pad=0.2)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def legend_image(segment_labels: List[str], colors: List[str] | None = None) -> io.BytesIO:
    """Render a legend strip (single row of colour patches + labels).

    Returns
    -------
    io.BytesIO  PNG image data.
    """
    if colors is None:
        colors = theme.LIKERT_COLORS
    n = len(segment_labels)
    fig, ax = plt.subplots(figsize=(5.5, 0.3))
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")
    ax.set_axis_off()

    patches = [
        mpatches.Patch(color=_hex_to_rgb(colors[i % len(colors)]),
                       label=segment_labels[i])
        for i in range(n)
    ]
    ax.legend(
        handles=patches,
        loc="center",
        ncol=n,
        fontsize=7,
        frameon=False,
        handlelength=1.2,
        handleheight=0.8,
        borderpad=0,
        columnspacing=1.0,
    )
    plt.tight_layout(pad=0)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
