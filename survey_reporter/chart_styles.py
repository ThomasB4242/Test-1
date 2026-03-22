"""Styled matplotlib chart renderers that return PNG BytesIO objects."""
from __future__ import annotations

import io
import math
import os
import re
import textwrap
from typing import List

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as _fm
# Register DM Sans if the TTF is installed, otherwise fall back to Liberation Sans
_DM_SANS_PATH = "/usr/local/share/fonts/dmsans/DMSans[opsz,wght].ttf"
if __import__("pathlib").Path(_DM_SANS_PATH).exists():
    _fm.fontManager.addfont(_DM_SANS_PATH)
    _DM_AVAILABLE = "DM Sans" in [f.name for f in _fm.fontManager.ttflist]
else:
    _DM_AVAILABLE = False
matplotlib.rcParams['font.family'] = 'DM Sans' if _DM_AVAILABLE else 'Liberation Sans'
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
AXIS_LEFT        = 0.38    # simple bar — short response labels
GRID_AXIS_LEFT   = 0.54    # grid / stacked — longer statement labels
AXIS_RIGHT       = 0.90
AXIS_TOP         = 0.95    # leave a sliver at top
AXIS_BOTTOM      = 0.03    # simple bar — no x-axis labels
GRID_AXIS_BOTTOM = 0.05    # grid — no x-axis labels needed

# Extra figure width (inches) added when a legend is present so all legend
# items fit on one row without being clipped.
LEGEND_EXTRA_W: float = 2.5


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
               axis_left: float, axis_bottom: float,
               fig_w: float = CHART_FIG_W) -> tuple[float, float]:
    """Return (r_x, r_y) in data coordinates for visually circular corners."""
    ax_w_in = (AXIS_RIGHT - axis_left) * fig_w
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
    fig_w: float | None = None,
) -> io.BytesIO:
    """Render a simple horizontal bar chart.

    Returns (buf, xlim_max) so callers can map data coords to slide coords.
    """
    if bar_colors is None:
        bar_colors = [color or theme.BAR_COLOR] * len(labels)
    if fig_h is None:
        fig_h = CHART_FIG_H
    if fig_w is None:
        fig_w = CHART_FIG_W

    n = len(labels)
    # Wrap long y-axis labels so they stay within the label area
    wrapped_labels = [_wrap_label(lbl, 22) for lbl in labels]

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("none")
    ax.set_facecolor("none")
    plt.subplots_adjust(
        left=AXIS_LEFT, right=AXIS_RIGHT,
        top=AXIS_TOP,   bottom=AXIS_BOTTOM,
    )

    max_val = max(values) if values else 1
    xlim_max = max_val * 1.12 + 3
    ax.set_xlim(0, xlim_max)

    # Rounded-corner radii scaled to give visually circular corners
    ax_w_in = (AXIS_RIGHT - AXIS_LEFT) * fig_w
    ax_h_in = (AXIS_TOP - AXIS_BOTTOM) * fig_h
    r_y = 0.60 * 0.38
    r_x = (r_y / (max(n, 1) / ax_h_in)) * (xlim_max / ax_w_in)

    for i, (label, val, clr) in enumerate(zip(labels, values, bar_colors)):
        _rounded_right_bar(ax, 0, val, i, 0.60, r_x, r_y, clr)
        if val >= 8:
            ax.text(val / 2, i, f"{int(round(val))}",
                    ha="center", va="center",
                    fontsize=14, color="white", fontweight="bold", zorder=4)
        else:
            ax.text(val + 0.8, i, f"{int(round(val))}",
                    ha="left", va="center",
                    fontsize=14, color="#1A1A1A", fontweight="bold", zorder=4)

    ax.set_yticks(list(range(n)))
    ax.set_yticklabels(wrapped_labels, fontsize=14, color="#1A1A1A",
                       linespacing=0.88)
    for lbl in ax.get_yticklabels():
        lbl.set_fontweight("bold")
    ax.set_ylim(-0.55, n - 0.45)
    ax.invert_yaxis()
    ax.tick_params(axis="y", length=0)
    ax.xaxis.set_visible(False)
    for name, spine in ax.spines.items():
        if name == "left":
            spine.set_visible(True)
            spine.set_linewidth(1.4)
            spine.set_color(theme.GRAY_DARK)
        else:
            spine.set_visible(False)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, transparent=True)
    plt.close(fig)
    buf.seek(0)
    # Attach xlim_max so callers can compute circle positions without re-deriving it
    buf._xlim_max = xlim_max  # type: ignore[attr-defined]
    return buf


# ---------------------------------------------------------------------------
# Pie chart
# ---------------------------------------------------------------------------

def render_pie(
    labels: List[str],
    values: List[float],
    colors: List[str] | None = None,
    fig_h: float = 4.0,
    fig_w: float = 4.0,
) -> io.BytesIO:
    """Render a pie chart.  Data labels inside slices are always white.
    Slices smaller than 3% get no label to avoid overlap clutter.
    """
    if colors is None:
        colors = list(theme.LIKERT_COLORS)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("none")
    ax.set_facecolor("none")

    _, _, autotexts = ax.pie(
        values,
        colors=[colors[i % len(colors)] for i in range(len(values))],
        autopct=lambda p: f"{p:.0f}" if p >= 3 else "",
        pctdistance=0.65,
        startangle=90,
        counterclock=False,
    )
    for t in autotexts:
        t.set_color("white")
        t.set_fontweight("bold")
        t.set_fontsize(13)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, transparent=True,
                bbox_inches="tight")
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
    fig_w: float | None = None,
    axis_left: float | None = None,
    axis_bottom: float | None = None,
    total_percents: List[float] | None = None,
    circle_color: str | None = None,
    legend_spec: List[tuple] | None = None,
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
    if fig_w is None:
        fig_w = CHART_FIG_W
    if axis_bottom is None:
        axis_bottom = GRID_AXIS_BOTTOM
    # Reserve extra bottom margin when embedding the legend.
    # Use a fixed absolute height (~0.70") so it doesn't eat into tall charts.
    if legend_spec:
        _legend_abs_h = 0.70  # inches
        axis_bottom = max(axis_bottom, _legend_abs_h / fig_h)

    n_rows = len(row_labels)
    wrapped_labels = [_wrap_label(lbl, 45) for lbl in row_labels]

    # Adaptive axis_left: scale with longest label so short-label grids get
    # wider bar area while long-label grids still have room to read labels.
    if axis_left is None:
        _max_chars = max(
            (max((len(part) for part in lbl.split('\n')), default=0)
             for lbl in wrapped_labels),
            default=15
        )
        # ~4 chars → 0.28, ~20 chars → 0.36, ~45 chars → 0.52; hard caps applied
        axis_left = max(0.26, min(GRID_AXIS_LEFT, 0.18 + _max_chars * 0.0075))

    # When a legend is present, widen the figure so all legend items fit on one
    # row.  Scale axis_left / AXIS_RIGHT proportionally so the bar area keeps
    # exactly the same absolute size — only the right-hand whitespace grows.
    if legend_spec:
        render_fig_w = fig_w + LEGEND_EXTRA_W
        _s  = fig_w / render_fig_w
        al  = axis_left * _s
        ar  = AXIS_RIGHT * _s
    else:
        render_fig_w = fig_w
        al  = axis_left
        ar  = AXIS_RIGHT

    fig, ax = plt.subplots(figsize=(render_fig_w, fig_h))
    fig.patch.set_facecolor("none")
    ax.set_facecolor("none")
    plt.subplots_adjust(
        left=al, right=ar,
        top=AXIS_TOP,   bottom=axis_bottom,
    )

    # Rounded corner radii for visually circular corners
    r_x, r_y = _bar_radii(fig_h, n_rows, al, axis_bottom, render_fig_w)

    # Bar height in matplotlib points — drives both segment-label and circle sizing
    _ax_h_in   = (AXIS_TOP - axis_bottom) * fig_h
    _bar_h_in  = 0.55 * _ax_h_in / max(n_rows, 1)
    _bar_h_pts = _bar_h_in * 72
    # Segment label font: proportional to bar height, capped at 14pt
    seg_label_fs = max(8, min(14, int(_bar_h_pts * 0.70)))

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
            if val >= 2:
                cx = lefts[row_i] + val / 2
                ax.text(cx, row_i, f"{int(round(val))}",
                        ha="center", va="center",
                        fontsize=seg_label_fs, color="white", fontweight="bold", zorder=4)
        lefts = [l + v for l, v in zip(lefts, vals)]

    # ── x-axis: grid lines only, NO tick labels ───────────────────────────────
    ax.set_xlim(0, 100)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xticklabels([])                          # no numbers
    ax.tick_params(axis="x", length=0)
    ax.xaxis.grid(True, linestyle="-", linewidth=0.6, color="#E0E0E0", zorder=0)
    ax.set_axisbelow(True)

    for name, spine in ax.spines.items():
        if name in ("bottom", "left"):
            spine.set_visible(True)
            spine.set_linewidth(1.4)
            spine.set_color(theme.GRAY_DARK)
        else:
            spine.set_visible(False)

    # ── y-axis: bold labels — font scales down for denser grids ──────────────
    # Calibrated so 3-row grid = 14pt, 13-row grid ≈ 12pt, min 10pt
    y_label_fs = max(10, min(14, 16 - n_rows // 3))
    ax.set_yticks(list(range(n_rows)))
    ax.set_yticklabels(wrapped_labels, fontsize=y_label_fs, color="#1A1A1A",
                       linespacing=0.88)
    for lbl in ax.get_yticklabels():
        lbl.set_fontweight("bold")
    ax.set_ylim(-0.55, n_rows - 0.45)
    ax.invert_yaxis()
    ax.tick_params(axis="y", length=0)

    # ── In-chart total circles ────────────────────────────────────────────────
    if total_percents:
        # Circle must be at least as large as the segment labels inside the bars.
        # Enforce: c_fs >= seg_label_fs, then back-derive minimum c_diam.
        c_diam   = max(seg_label_fs / 0.38, _bar_h_pts, 16)
        c_fs     = max(seg_label_fs, int(c_diam * 0.38))
        border_w = max(1.5, c_diam * 0.12)
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
                    ha="center", va="center_baseline",
                    fontsize=c_fs, color="#1A1A1A", fontweight="bold",
                    zorder=6, clip_on=False)

    # ── Embedded legend (drawn inside the figure below the axes) ─────────────
    if legend_spec:
        handles = []
        for lbl, clr, is_circle in legend_spec:
            if is_circle:
                cclr = _hex_to_rgb(clr)
                h = mlines.Line2D(
                    [], [], marker="o", color="none",
                    markerfacecolor="white", markeredgecolor=cclr,
                    markeredgewidth=1.5, markersize=9, label=lbl,
                )
            else:
                h = mpatches.Patch(facecolor=_hex_to_rgb(clr), label=lbl)
            handles.append(h)
        ax_center = (al + ar) / 2
        # All items on one row; figure is widened by LEGEND_EXTRA_W so no clipping
        fig.legend(
            handles=handles,
            loc="lower center",
            bbox_to_anchor=(ax_center, 0.01),
            bbox_transform=fig.transFigure,
            ncol=len(handles),
            fontsize=14,
            labelcolor="#1A1A1A",
            frameon=False,
            handlelength=1.2,
            handleheight=0.9,
            borderpad=0,
            columnspacing=0.7,
        )

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, transparent=True)
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
            teal_rgb = _hex_to_rgb(theme.TEAL_MID)
            h = mlines.Line2D(
                [], [], color="none", marker="o",
                markerfacecolor="white", markeredgecolor=teal_rgb,
                markeredgewidth=1.5, markersize=9, label=label,
            )
        else:
            h = mpatches.Patch(color=clr, label=label)
        handles.append(h)

    fig, ax = plt.subplots(figsize=(fig_w, 0.30))
    fig.patch.set_alpha(0.0)
    ax.set_facecolor("none")
    ax.set_axis_off()
    ax.legend(handles=handles, loc="center", ncol=n, fontsize=9,
              frameon=False, handlelength=1.2, handleheight=0.8,
              borderpad=0, columnspacing=0.9)
    plt.tight_layout(pad=0)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI,
                bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Pie chart
# ---------------------------------------------------------------------------

_PIE_DARK_FILLS = {theme.TEAL_DARK, theme.TEAL_MID, theme.CORAL, theme.SALMON}


def render_pie_chart(
    labels: List[str],
    values: List[float],
    colors: List[str] | None = None,
    fig_h: float | None = None,
    fig_w: float | None = None,
) -> io.BytesIO:
    """Pie chart with % numbers inside segments and labelled leader lines outside."""
    if colors is None:
        colors = (theme.LIKERT_COLORS * 4)[:len(labels)]
    if fig_h is None:
        fig_h = CHART_FIG_H
    if fig_w is None:
        fig_w = CHART_FIG_W

    total = sum(values) or 1.0
    rgb_colors = [_hex_to_rgb(c) for c in colors]

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("none")
    ax.set_facecolor("none")

    # Draw wedges only (no built-in labels / autopct — we draw them manually)
    wedges, _ = ax.pie(
        values,
        colors=rgb_colors,
        startangle=90,
        wedgeprops={"linewidth": 2, "edgecolor": "white"},
        radius=0.75,
    )

    # Numbers inside each segment
    for wedge, val, clr in zip(wedges, values, colors):
        mid_angle = math.radians((wedge.theta1 + wedge.theta2) / 2)
        rx, ry = 0.46 * math.cos(mid_angle), 0.46 * math.sin(mid_angle)
        pct = int(round(val / total * 100))
        txt_clr = "white" if clr in _PIE_DARK_FILLS else theme.GRAY_DARK
        ax.text(rx, ry, f"{pct}%",
                ha="center", va="center",
                fontsize=12, fontweight="bold", color=txt_clr, zorder=5)

    # External labels with pointer lines
    for wedge, label in zip(wedges, labels):
        mid_angle = math.radians((wedge.theta1 + wedge.theta2) / 2)
        cos_a, sin_a = math.cos(mid_angle), math.sin(mid_angle)
        # Line: from wedge edge to an elbow just outside
        x0, y0 = 0.78 * cos_a, 0.78 * sin_a   # wedge surface
        x1, y1 = 0.98 * cos_a, 0.98 * sin_a   # elbow
        ax.plot([x0, x1], [y0, y1],
                color=theme.GRAY_DARK, lw=0.9, clip_on=False)
        ha = "left" if cos_a >= 0 else "right"
        x_txt = x1 + (0.04 if ha == "left" else -0.04)
        ax.text(x_txt, y1, label,
                ha=ha, va="center",
                fontsize=14, fontweight="bold", color="#1A1A1A",
                clip_on=False)

    ax.set_xlim(-1.55, 1.55)
    ax.set_ylim(-1.25, 1.25)
    ax.set_aspect("equal")

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, transparent=True,
                bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Rounded-top-corner bar primitive  (for vertical / column charts)
# ---------------------------------------------------------------------------

def _rounded_top_bar(ax, x_center: float, y0: float, y1: float,
                     bar_w: float, r_x: float, r_y: float,
                     hex_color: str, zorder: int = 3) -> None:
    """Draw a vertical bar with rounded top corners, sharp bottom corners."""
    if y1 <= y0:
        return
    r_x = min(r_x, bar_w * 0.48)
    r_y = min(r_y, (y1 - y0) * 0.48)
    x0 = x_center - bar_w / 2
    x1 = x_center + bar_w / 2
    verts = [
        (x0,       y0),
        (x0,       y1 - r_y),
        (x0,       y1),           # ctrl – top-left corner
        (x0 + r_x, y1),
        (x1 - r_x, y1),
        (x1,       y1),           # ctrl – top-right corner
        (x1,       y1 - r_y),
        (x1,       y0),
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


# ---------------------------------------------------------------------------
# Stacked column chart  (vertical stacked bars — for tracking data)
# ---------------------------------------------------------------------------

def render_stacked_column(
    col_labels: List[str],
    segments: List[dict],
    colors: List[str] | None = None,
    fig_h: float | None = None,
    fig_w: float | None = None,
    total_percents: List[float] | None = None,
    circle_color: str | None = None,
    legend_spec: List[tuple] | None = None,
) -> io.BytesIO:
    """Vertical stacked bar chart (columns) for tracking data.

    segments : [{label, values}] where values[i] = % for column i.
    total_percents : top-box total per column — circle drawn ON the bar at
                     y = total_pct (same principle as horizontal bar circles).
    Legend drawn on the right side of the axes.
    """
    n_cols = len(col_labels)
    if colors is None:
        colors = list(theme.LIKERT_COLORS)
    if fig_h is None:
        fig_h = 4.5
    if fig_w is None:
        fig_w = max(5.0, n_cols * 1.4)

    # Margins: leave right room for legend
    AL = 0.07   # axis left
    AR = 0.72   # axis right (legend goes 0.73–1.0)
    AT = 0.92
    AB = 0.12   # bottom for x-tick labels

    bar_w = 0.65   # bar width in data units (x-axis = 0..n_cols-1)
    r_x_data = bar_w * 0.22

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("none")
    ax.set_facecolor("none")
    plt.subplots_adjust(left=AL, right=AR, top=AT, bottom=AB)

    # Pre-compute topmost non-zero segment per column
    last_seg_idx = [-1] * n_cols
    for i, seg in enumerate(segments):
        for col_i, val in enumerate(seg["values"]):
            if val > 0:
                last_seg_idx[col_i] = i

    # Draw segments bottom-to-top
    bottoms = [0.0] * n_cols
    for i, seg in enumerate(segments):
        clr = colors[i % len(colors)]
        for col_i, val in enumerate(seg["values"]):
            if val <= 0:
                continue
            y0 = bottoms[col_i]
            y1 = y0 + val
            is_top = (i == last_seg_idx[col_i])
            ry = (val * 0.18) if is_top else 0.0
            _rounded_top_bar(ax, col_i, y0, y1, bar_w, r_x_data, ry, clr)
            if val >= 2:
                ax.text(col_i, y0 + val / 2, f"{int(round(val))}",
                        ha="center", va="center_baseline",
                        fontsize=10, color="white", fontweight="bold", zorder=4)
            bottoms[col_i] = y1

    # Total circles — at (col_i, total_pct) matching the percentage on the y-axis
    if total_percents:
        cclr_hex = circle_color or theme.TEAL_MID
        cclr = _hex_to_rgb(cclr_hex)
        c_diam = 22.0
        c_fs   = max(8, int(c_diam * 0.38))
        border_w = max(1.5, c_diam * 0.12)
        for col_i, tp in enumerate(total_percents):
            if tp <= 0:
                continue
            ax.plot(col_i, tp, "o",
                    markersize=c_diam, color="white",
                    markeredgecolor=cclr, markeredgewidth=border_w,
                    zorder=5, clip_on=False)
            ax.text(col_i, tp, str(int(round(tp))),
                    ha="center", va="center_baseline",
                    fontsize=c_fs, color="#1A1A1A", fontweight="bold",
                    zorder=6, clip_on=False)

    # Axes styling
    ax.set_xlim(-0.55, n_cols - 0.45)
    ax.set_ylim(0, 100)
    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(col_labels, fontsize=12, fontweight="bold", color="#1A1A1A")
    ax.tick_params(axis="x", length=0)
    ax.yaxis.set_visible(False)
    ax.yaxis.grid(True, linestyle="-", linewidth=0.6, color="#E0E0E0", zorder=0)
    ax.set_axisbelow(True)
    for name, spine in ax.spines.items():
        if name == "bottom":
            spine.set_visible(True)
            spine.set_linewidth(1.4)
            spine.set_color(theme.GRAY_DARK)
        else:
            spine.set_visible(False)

    # Right-side legend — reversed so topmost segment appears at top of legend
    if legend_spec:
        legend_x = AR + 0.04
        legend_items = list(reversed(legend_spec))
        n_items = len(legend_items)
        legend_y_start = AT
        legend_y_step  = (AT - AB) / max(n_items, 1) * 0.88
        for k, (lbl, clr, is_circle) in enumerate(legend_items):
            ly = legend_y_start - k * legend_y_step
            if is_circle:
                circ = mpatches.Circle(
                    (legend_x + 0.014, ly - 0.013), radius=0.013,
                    facecolor="white", edgecolor=_hex_to_rgb(clr),
                    linewidth=1.5, transform=fig.transFigure,
                    clip_on=False, zorder=10,
                )
                fig.add_artist(circ)
            else:
                sq = mpatches.FancyBboxPatch(
                    (legend_x, ly - 0.026), 0.028, 0.022,
                    boxstyle="square,pad=0",
                    facecolor=_hex_to_rgb(clr), edgecolor="none",
                    transform=fig.transFigure, clip_on=False, zorder=10,
                )
                fig.add_artist(sq)
            fig.text(legend_x + 0.035, ly - 0.007, lbl,
                     ha="left", va="top", fontsize=8.5,
                     color="#1A1A1A", transform=fig.transFigure)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Clustered column chart  (grouped bars — e.g. statements × age groups)
# ---------------------------------------------------------------------------

def render_cluster_column(
    cluster_labels: List[str],
    series: List[dict],
    colors: List[str] | None = None,
    fig_h: float | None = None,
    fig_w: float | None = None,
    legend_spec: List[tuple] | None = None,
) -> io.BytesIO:
    """Grouped vertical bar chart.

    series : [{label, values}] — one per series (e.g. age group).
              values[i] = % for cluster i.
    Legend drawn along the bottom (one entry per series).
    """
    import numpy as np

    n_clusters = len(cluster_labels)
    n_series   = len(series)
    if colors is None:
        colors = list(theme.CLUSTER_COLORS)
    if fig_h is None:
        fig_h = 4.5
    if fig_w is None:
        fig_w = max(6.0, n_clusters * 1.2)

    AL = 0.06
    AR = 0.97
    AT = 0.93
    AB = 0.22   # extra bottom margin for x-labels + legend

    bar_w     = 0.16
    group_gap = 0.08
    total_group_w = n_series * bar_w + group_gap
    import numpy as np_inner
    x = np_inner.arange(n_clusters) * (total_group_w + 0.12)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("none")
    ax.set_facecolor("none")
    plt.subplots_adjust(left=AL, right=AR, top=AT, bottom=AB)

    for s_i, ser in enumerate(series):
        clr  = colors[s_i % len(colors)]
        xs   = x + s_i * bar_w - (n_series - 1) * bar_w / 2
        vals = [float(v) for v in ser["values"]]
        ax.bar(xs, vals, width=bar_w, color=_hex_to_rgb(clr),
               zorder=3, edgecolor="none")
        for xi, vi in zip(xs, vals):
            if vi >= 8:
                ax.text(xi, vi / 2, f"{int(round(vi))}",
                        ha="center", va="center_baseline",
                        fontsize=8, color="white", fontweight="bold", zorder=4)

    # Axes
    ax.set_ylim(0, 100)
    ax.set_xticks(x)
    wrapped = [_wrap_label(lbl, 14) for lbl in cluster_labels]
    ax.set_xticklabels(wrapped, fontsize=10, fontweight="bold",
                       color="#1A1A1A", linespacing=0.9)
    ax.tick_params(axis="x", length=0)
    ax.yaxis.set_visible(False)
    ax.yaxis.grid(True, linestyle="-", linewidth=0.6, color="#E0E0E0", zorder=0)
    ax.set_axisbelow(True)
    for name, spine in ax.spines.items():
        if name == "bottom":
            spine.set_visible(True)
            spine.set_linewidth(1.4)
            spine.set_color(theme.GRAY_DARK)
        else:
            spine.set_visible(False)

    # Bottom legend — one entry per series, left-to-right
    if legend_spec:
        handles = [mpatches.Patch(color=_hex_to_rgb(clr), label=lbl)
                   for lbl, clr, _ in legend_spec]
        ax.legend(handles=handles, loc="upper center",
                  bbox_to_anchor=(0.5, -0.18),
                  ncol=n_series, fontsize=9, frameon=False,
                  handlelength=1.0, handleheight=0.8,
                  columnspacing=1.0, borderpad=0)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, transparent=True,
                bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf
