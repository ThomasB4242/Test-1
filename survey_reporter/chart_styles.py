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
from matplotlib.path import Path
from matplotlib.patches import PathPatch

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
GRID_AXIS_BOTTOM = 0.05    # grid — no x-axis labels needed (was 0.12)


def _hex_to_rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


def _text_color_for_bg(hex_bg: str) -> str:
    """White on dark/medium backgrounds, gray on light ones."""
    return "white" if hex_bg in {theme.TEAL_DARK, theme.TEAL_MID, theme.CORAL} else theme.GRAY_DARK


def _wrap_label(text: str, max_chars: int = 32) -> str:
    return "\n".join(textwrap.wrap(text, max_chars))


# ---------------------------------------------------------------------------
# Rounded-right-corner bar primitive
# ---------------------------------------------------------------------------

def _rounded_right_bar(ax, x0: float, x1: float, y_center: float,
                        bar_h: float, r_x: float, r_y: float,
                        hex_color: str, zorder: int = 3) -> None:
    """Draw a horizontal bar with rounded right corners, sharp left corners."""
    if x1 <= x0:
        return
    r_x = min(r_x, (x1 - x0) * 0.48)
    r_y = min(r_y, bar_h * 0.48)
    y0 = y_center - bar_h / 2
    y1 = y_center + bar_h / 2
    verts = [
        (x0,       y0),
        (x1 - r_x, y0),
        (x1,       y0),           # ctrl – bottom-right corner
        (x1,       y0 + r_y),
        (x1,       y1 - r_y),
        (x1,       y1),           # ctrl – top-right corner
        (x1 - r_x, y1),
        (x0,       y1),
        (x0,       y0),
    ]
    codes = [
        Path.MOVETO,
        Path.LINETO,
        Path.CURVE3, Path.CURVE3,
        Path.LINETO,
        Path.CURVE3, Path.CURVE3,
        Path.LINETO,
        Path.CLOSEPOLY,
    ]
    fc = _hex_to_rgb(hex_color)
    ax.add_patch(PathPatch(Path(verts, codes),
                           facecolor=fc, edgecolor="none", zorder=zorder))


def _bar_radii(fig_h: float, n_rows: int,
               axis_left: float, axis_bottom: float) -> tuple[float, float]:
    """Return (r_x, r_y) in data coordinates for visually circular corners."""
    ax_w_in = (AXIS_RIGHT - axis_left) * CHART_FIG_W
    ax_h_in = (AXIS_TOP - axis_bottom) * fig_h
    r_y = 0.55 * 0.22                                  # 22% of bar height in y-data-units
    r_x = (r_y / (max(n_rows, 1) / ax_h_in)) * (100.0 / ax_w_in)
    return r_x, r_y


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
    """Render a simple horizontal bar chart."""
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
    xlim_max = max_val * 1.12 + 3
    ax.set_xlim(0, xlim_max)

    # Rounded-corner radii scaled to give visually circular corners
    ax_w_in = (AXIS_RIGHT - AXIS_LEFT) * CHART_FIG_W
    ax_h_in = (AXIS_TOP - AXIS_BOTTOM) * fig_h
    r_y = 0.60 * 0.38
    r_x = (r_y / (max(n, 1) / ax_h_in)) * (xlim_max / ax_w_in)

    for i, (label, val, clr) in enumerate(zip(labels, values, bar_colors)):
        _rounded_right_bar(ax, 0, val, i, 0.60, r_x, r_y, clr)
        if val >= 8:
            ax.text(val / 2, i, f"{int(round(val))}",
                    ha="center", va="center",
                    fontsize=10, color="white", fontweight="bold", zorder=4)
        else:
            ax.text(val + 0.8, i, f"{int(round(val))}",
                    ha="left", va="center",
                    fontsize=10, color=theme.GRAY_DARK, fontweight="bold", zorder=4)

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
    total_percents: List[float] | None = None,
    circle_color: str | None = None,
) -> io.BytesIO:
    """Render a stacked 0–100 horizontal bar chart for grid questions.

    Parameters
    ----------
    row_labels      : One label per row, index 0 at the TOP.
    segments        : List of ``{label, values}`` dicts (one per scale option).
    colors          : Hex colour per segment.
    fig_h           : Figure height (inches). Defaults to CHART_FIG_H.
    axis_left       : Left subplot margin fraction. Defaults to GRID_AXIS_LEFT.
    axis_bottom     : Bottom subplot margin fraction. Defaults to GRID_AXIS_BOTTOM.
    total_percents  : Per-row top-box total (0–100). When supplied a filled
                      circle is drawn ON the chart at each total position.
    circle_color    : Hex colour for the total circles. Defaults to TEAL_MID.
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

    # Rounded corner radii for visually circular corners
    r_x, r_y = _bar_radii(fig_h, n_rows, axis_left, axis_bottom)

    # Pre-compute which segment is the last (rightmost) non-zero for each row
    last_seg_idx = [-1] * n_rows
    for i, seg in enumerate(segments):
        for row_i, val in enumerate(seg["values"]):
            if val > 0:
                last_seg_idx[row_i] = i

    # Draw stacked segments — only the last non-zero segment gets rounded right corners
    lefts = [0.0] * n_rows
    for i, seg in enumerate(segments):
        clr  = colors[i % len(colors)]
        vals = seg["values"]
        txt_clr = _text_color_for_bg(clr)
        for row_i, val in enumerate(vals):
            if val <= 0:
                continue
            is_last = (i == last_seg_idx[row_i])
            _rounded_right_bar(ax,
                                lefts[row_i], lefts[row_i] + val,
                                row_i, 0.55,
                                r_x if is_last else 0.0,
                                r_y if is_last else 0.0,
                                clr)
            if val >= 7:
                cx = lefts[row_i] + val / 2
                ax.text(cx, row_i, f"{int(round(val))}",
                        ha="center", va="center",
                        fontsize=9, color=txt_clr, fontweight="bold", zorder=4)
        lefts = [l + v for l, v in zip(lefts, vals)]

    # ── x-axis: grid lines only, NO tick labels ───────────────────────────────
    ax.set_xlim(0, 100)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xticklabels([])                          # no numbers
    ax.tick_params(axis="x", length=0)
    ax.xaxis.grid(True, linestyle="-", linewidth=0.6, color="#E0E0E0", zorder=0)
    ax.set_axisbelow(True)

    for name, spine in ax.spines.items():
        if name == "bottom":
            spine.set_visible(True)
            spine.set_linewidth(1.4)
            spine.set_color(theme.GRAY_DARK)
        else:
            spine.set_visible(False)

    # ── y-axis: bold labels ───────────────────────────────────────────────────
    ax.set_yticks(list(range(n_rows)))
    ax.set_yticklabels(wrapped_labels, fontsize=label_fs, color=theme.GRAY_DARK,
                       linespacing=1.1, fontweight="bold")
    ax.set_ylim(-0.55, n_rows - 0.45)
    ax.invert_yaxis()
    ax.tick_params(axis="y", length=0)

    # ── In-chart total circles ────────────────────────────────────────────────
    if total_percents:
        # Diameter in matplotlib points ≈ bar height in inches × 72 pt/in
        ax_h_in  = (AXIS_TOP - axis_bottom) * fig_h
        bar_h_in = 0.55 * ax_h_in / max(n_rows, 1)
        c_diam   = bar_h_in * 72          # points
        c_fs     = max(6, int(c_diam * 0.38))
        border_w = max(1.5, c_diam * 0.12)   # border thickness scales with size
        cclr_hex = circle_color or theme.TEAL_MID
        cclr     = _hex_to_rgb(cclr_hex)
        for row_i, tp in enumerate(total_percents):
            if tp <= 0:
                continue
            # White fill with thick teal border
            ax.plot(tp, row_i, "o",
                    markersize=c_diam,
                    color="white",
                    markeredgecolor=cclr,
                    markeredgewidth=border_w,
                    zorder=5, clip_on=False)
            ax.text(tp, row_i, str(int(round(tp))),
                    ha="center", va="center",
                    fontsize=c_fs, color=cclr, fontweight="bold",
                    zorder=6, clip_on=False)

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
    """Render a small horizontal legend strip."""
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
