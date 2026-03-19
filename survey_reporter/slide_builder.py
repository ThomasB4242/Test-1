"""Assemble a PowerPoint presentation from QuestionResult objects and YAML annotations."""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from . import theme
from . import chart_styles
from .analyzer import QuestionResult


# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------

def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------------------------------------------------------------------
# Bar colour assignment
# ---------------------------------------------------------------------------

# Ordered palette for scale questions:
#   top-box bars (most → least positive):  TEAL_DARK, TEAL_MID
#   neutral bar:                            GRAY_MID
#   bottom-box bars (least → most negative): SALMON, CORAL
_TOP_COLORS = [theme.TEAL_DARK, theme.TEAL_MID, theme.TEAL_LIGHT]
_BOT_COLORS = [theme.SALMON, theme.CORAL]


def _scale_bar_colors(labels_t2b: list[str],
                      top_box_spec: dict | None,
                      bottom_box_spec: dict | None) -> list[str]:
    """Return per-bar hex colours based on top/bottom box membership.

    The first top-box bar encountered (most positive, i.e. index 0 in
    *labels_t2b*) gets TEAL_DARK; subsequent top-box bars step through
    _TOP_COLORS.  Similarly the first bottom-box bar gets SALMON, the
    next CORAL.  Everything else gets GRAY_MID.
    """
    top_vals = set(top_box_spec.get("values", []) if top_box_spec else [])
    bot_vals = set(bottom_box_spec.get("values", []) if bottom_box_spec else [])

    if not top_vals and not bot_vals:
        return [theme.BAR_COLOR] * len(labels_t2b)

    top_idx = bot_idx = 0
    colors: list[str] = []
    for label in labels_t2b:
        if label in top_vals:
            colors.append(_TOP_COLORS[min(top_idx, len(_TOP_COLORS) - 1)])
            top_idx += 1
        elif label in bot_vals:
            colors.append(_BOT_COLORS[min(bot_idx, len(_BOT_COLORS) - 1)])
            bot_idx += 1
        else:
            colors.append(theme.GRAY_MID)
    return colors


# ---------------------------------------------------------------------------
# Circle positioning helpers
# ---------------------------------------------------------------------------

def _bar_circle_center_y(label_index: float, n_bars: int,
                         chart_top: float, chart_h: float) -> float:
    """Return slide y-coordinate (inches) for the center of bar at *label_index*.

    Uses exported subplots_adjust margins from chart_styles so the circle
    aligns with the actual rendered bar.  With invert_yaxis, index 0 is TOP.
    """
    ax_top  = chart_top + (1.0 - chart_styles.AXIS_TOP) * chart_h
    ax_h    = (chart_styles.AXIS_TOP - chart_styles.AXIS_BOTTOM) * chart_h
    return ax_top + (label_index + 0.5) * (ax_h / n_bars)


def _circle_cx(chart_left: float, chart_w: float, radius: float) -> float:
    """Place circle so it overlaps the right end of the chart area."""
    return chart_left + chart_w - radius * 0.5


# ---------------------------------------------------------------------------
# pptx primitive helpers
# ---------------------------------------------------------------------------

def _add_textbox(slide, ltwh, text: str, font_size: int, color: str,
                 bold: bool = False, italic: bool = False,
                 align=PP_ALIGN.LEFT, word_wrap: bool = True) -> Any:
    l, t, w, h = ltwh
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = _rgb(color)
    run.font.bold = bold
    run.font.italic = italic
    return txBox


def _add_multiline_textbox(slide, ltwh, lines: list[str], font_size: int,
                           color: str, bold: bool = False) -> Any:
    l, t, w, h = ltwh
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = _rgb(color)
        run.font.bold = bold
    return txBox


def _add_filled_rect(slide, ltwh, fill_color: str) -> Any:
    l, t, w, h = ltwh
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_color)
    shape.line.fill.background()
    return shape


def _add_oval(slide, ltwh, fill_color: str) -> Any:
    l, t, w, h = ltwh
    shape = slide.shapes.add_shape(9, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_color)
    shape.line.fill.background()
    return shape


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ---------------------------------------------------------------------------
# Shared slide furniture
# ---------------------------------------------------------------------------

def _add_header_stripe(slide):
    _add_filled_rect(slide, (0, 0, theme.SLIDE_W, 0.50), theme.TEAL_DARK)


def _add_footer(slide, base_note: str, page_num: int | None):
    _add_textbox(slide, theme.FOOTER_BOX, base_note,
                 font_size=theme.FONT_FOOTER, color=theme.GRAY_DARK, italic=True)
    if page_num is not None:
        _add_textbox(slide, theme.PAGE_NUM_BOX, str(page_num),
                     font_size=theme.FONT_FOOTER, color=theme.WHITE,
                     align=PP_ALIGN.RIGHT)


def _add_logo(slide, logo_path: str | None):
    if logo_path and Path(logo_path).exists():
        try:
            slide.shapes.add_picture(logo_path,
                                     Inches(theme.LOGO_BOX[0]), Inches(theme.LOGO_BOX[1]),
                                     height=Inches(theme.LOGO_BOX[3]))
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Circle badge
# ---------------------------------------------------------------------------

def _draw_circle_badge(slide, cx: float, cy: float, r: float,
                       fill_color: str, pct: int, label: str):
    """Draw a filled circle centred at (cx, cy) with radius r (all inches)."""
    l, t, d = cx - r, cy - r, r * 2
    _add_oval(slide, (l, t, d, d), fill_color)
    _add_textbox(slide, (l, t + d * 0.10, d, d * 0.52),
                 f"{pct}%",
                 font_size=theme.FONT_CIRCLE, color=theme.WHITE,
                 bold=True, align=PP_ALIGN.CENTER)
    _add_textbox(slide, (l, t + d * 0.58, d, d * 0.36),
                 label,
                 font_size=theme.FONT_CIRCLE_SM, color=theme.WHITE,
                 align=PP_ALIGN.CENTER)


def _sum_box(freq_by_label: dict, values: list[str]) -> int:
    return int(round(sum(freq_by_label.get(v, 0.0) for v in values)))


# ---------------------------------------------------------------------------
# Cover / section / commentary slides
# ---------------------------------------------------------------------------

def add_cover_slide(prs: Presentation, spec: dict, logo_path: str | None = None):
    slide = _blank_slide(prs)
    _add_filled_rect(slide, (0, 0, theme.SLIDE_W, theme.SLIDE_H), theme.TEAL_DARK)
    title    = spec.get("title", "Survey Report")
    subtitle = spec.get("subtitle", "")
    base     = spec.get("base", "")
    _add_textbox(slide, (0.5, 2.5, 9.0, 1.4), title,
                 font_size=32, color=theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
    if subtitle:
        _add_textbox(slide, (0.5, 3.9, 9.0, 0.7), subtitle,
                     font_size=18, color=theme.TEAL_LIGHT, align=PP_ALIGN.CENTER)
    if base:
        _add_textbox(slide, (0.5, 6.8, 9.0, 0.4), base,
                     font_size=10, color=theme.TEAL_LIGHT, align=PP_ALIGN.CENTER)
    if logo_path and Path(logo_path).exists():
        try:
            slide.shapes.add_picture(logo_path, Inches(0.25), Inches(0.15),
                                     height=Inches(0.35))
        except Exception:
            pass


def add_section_slide(prs: Presentation, spec: dict, logo_path: str | None = None):
    slide = _blank_slide(prs)
    _add_filled_rect(slide, (0, 0, theme.SLIDE_W, theme.SLIDE_H), theme.TEAL_DARK)
    title = spec.get("title", "Section")
    _add_textbox(slide, (0.5, 3.0, 9.0, 1.2), title,
                 font_size=28, color=theme.WHITE, bold=True, align=PP_ALIGN.LEFT)


def add_commentary_slide(prs: Presentation, spec: dict, page_num: int | None = None):
    slide = _blank_slide(prs)
    _add_header_stripe(slide)
    heading = spec.get("heading", "Commentary")
    _add_textbox(slide, theme.HEADING_BOX, heading,
                 font_size=theme.FONT_HEADING, color=theme.CORAL, bold=True)
    lines: list[str] = []
    bold_intro = spec.get("bold_intro", "")
    if bold_intro:
        lines.append(bold_intro)
    for b in spec.get("bullets", []):
        lines.append(f"• {b}" if not b.startswith("•") else b)
    if lines:
        _add_multiline_textbox(slide, (0.30, 1.30, 9.40, 5.80), lines,
                                font_size=theme.FONT_BODY, color=theme.GRAY_DARK)
    _add_footer(slide, spec.get("base", ""), page_num)


# ---------------------------------------------------------------------------
# Chart slide  (single question — individual bars)
# ---------------------------------------------------------------------------

def add_chart_slide(
    prs: Presentation,
    spec: dict,
    result: QuestionResult | None,
    page_num: int | None = None,
    logo_path: str | None = None,
):
    slide = _blank_slide(prs)
    _add_header_stripe(slide)
    _add_logo(slide, logo_path)

    heading         = spec.get("heading", result.label if result else "")
    question        = spec.get("question", "")
    chart_type      = spec.get("chart_type", "bar")
    annot_lines     = spec.get("annotations", [])
    base_note       = spec.get("base", "")
    top_box_spec    = spec.get("top_box")
    bottom_box_spec = spec.get("bottom_box")

    _add_textbox(slide, theme.HEADING_BOX, heading,
                 font_size=theme.FONT_HEADING, color=theme.CORAL, bold=True)
    if question:
        _add_textbox(slide, theme.QUESTION_BOX, question,
                     font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    # ── Build chart ──────────────────────────────────────────────────────────
    chart_buf  = None
    n_bars     = 0
    labels_t2b = []

    if result is not None and result.type == "categorical" and result.frequencies:
        freqs = result.frequencies

        if chart_type == "stacked":
            row_labels = [""]
            segments = [{"label": f.label, "values": [f.percent]} for f in freqs]
            colors   = theme.LIKERT_COLORS[:len(segments)]
            chart_buf = chart_styles.render_stacked_bar(row_labels, segments, colors=colors)
            n_bars = 1
            labels_t2b = [""]
        else:
            # "bar" — reversed SPSS code order → most positive at top
            labels_t2b   = list(reversed([f.label   for f in freqs]))
            percents_t2b = list(reversed([f.percent for f in freqs]))
            n_bars       = len(labels_t2b)
            bar_colors   = _scale_bar_colors(labels_t2b, top_box_spec, bottom_box_spec)
            chart_buf    = chart_styles.render_simple_bar(
                labels_t2b, percents_t2b, bar_colors=bar_colors
            )

    cl, ct, cw, ch = theme.CHART_BOX
    if chart_buf is not None:
        chart_buf.seek(0)
        slide.shapes.add_picture(chart_buf,
                                 Inches(cl), Inches(ct),
                                 width=Inches(cw), height=Inches(ch))

    # ── Circle badges ────────────────────────────────────────────────────────
    if result is not None and result.type == "categorical":
        freq_by_label = {f.label: f.percent for f in result.frequencies}
        r = 0.72
        cx = _circle_cx(cl, cw, r)

        if chart_type == "bar" and n_bars > 0:
            if top_box_spec:
                idxs = [i for i, l in enumerate(labels_t2b)
                        if l in set(top_box_spec.get("values", []))]
                if idxs:
                    cy = _bar_circle_center_y(
                        sum(idxs) / len(idxs), n_bars, ct, ch)
                    pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
                    _draw_circle_badge(slide, cx, cy, r, theme.TEAL_MID,
                                       pct, top_box_spec.get("label", ""))

            if bottom_box_spec:
                idxs = [i for i, l in enumerate(labels_t2b)
                        if l in set(bottom_box_spec.get("values", []))]
                if idxs:
                    cy = _bar_circle_center_y(
                        sum(idxs) / len(idxs), n_bars, ct, ch)
                    pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
                    _draw_circle_badge(slide, cx, cy, r * 0.90, theme.CORAL,
                                       pct, bottom_box_spec.get("label", ""))
        else:
            cx_s, cy_s, r_s = theme.TOP_CIRCLE_CENTER_R
            if top_box_spec:
                pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
                _draw_circle_badge(slide, cx_s, cy_s, r_s, theme.TEAL_MID,
                                   pct, top_box_spec.get("label", ""))
            cx_b, cy_b, r_b = theme.BOTTOM_CIRCLE_CENTER_R
            if bottom_box_spec:
                pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
                _draw_circle_badge(slide, cx_b, cy_b, r_b, theme.CORAL,
                                   pct, bottom_box_spec.get("label", ""))

    # ── Annotation panel ─────────────────────────────────────────────────────
    if annot_lines:
        _add_multiline_textbox(slide, theme.ANNOT_BOX, annot_lines,
                                font_size=theme.FONT_BODY, color=theme.GRAY_DARK)

    _add_footer(slide, base_note, page_num)


# ---------------------------------------------------------------------------
# Grid slide  (multiple questions → stacked bar, one row per question)
# ---------------------------------------------------------------------------

def add_grid_slide(
    prs: Presentation,
    spec: dict,
    results: list[QuestionResult],
    page_num: int | None = None,
    logo_path: str | None = None,
):
    """Render multiple related questions as a stacked horizontal bar chart.

    Each row in the chart = one question / item.
    Each colour segment = one response option on the shared scale.

    YAML spec fields:
        variables    : list of variable names (already resolved to QuestionResult objects)
        scale_order  : ordered list of response labels (most-positive first)
        top_box      : {label, values}
        bottom_box   : {label, values}
        row_labels   : optional list of short row labels (overrides SPSS variable labels)
    """
    slide = _blank_slide(prs)
    _add_header_stripe(slide)
    _add_logo(slide, logo_path)

    heading         = spec.get("heading", "")
    question        = spec.get("question", "")
    annot_lines     = spec.get("annotations", [])
    base_note       = spec.get("base", "")
    top_box_spec    = spec.get("top_box")
    bottom_box_spec = spec.get("bottom_box")
    scale_order     = spec.get("scale_order", [])
    row_label_overrides = spec.get("row_labels", [])

    _add_textbox(slide, theme.HEADING_BOX, heading,
                 font_size=theme.FONT_HEADING, color=theme.CORAL, bold=True)
    if question:
        _add_textbox(slide, theme.QUESTION_BOX, question,
                     font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    if not results:
        _add_footer(slide, base_note, page_num)
        return

    # Determine scale order from first result if not specified
    if not scale_order and results:
        scale_order = [f.label for f in results[0].frequencies]

    # Row labels: use overrides if provided, else SPSS variable labels
    row_labels = []
    for i, result in enumerate(results):
        if i < len(row_label_overrides) and row_label_overrides[i]:
            row_labels.append(row_label_overrides[i])
        else:
            # Truncate long SPSS labels for readability
            lbl = result.label
            if len(lbl) > 40:
                lbl = lbl[:37] + "…"
            row_labels.append(lbl)

    # Build segments (one per scale option in order)
    segments: list[dict] = []
    for opt_label in scale_order:
        values = []
        for result in results:
            pct = next((f.percent for f in result.frequencies
                        if f.label == opt_label), 0.0)
            values.append(pct)
        segments.append({"label": opt_label, "values": values})

    # Assign colours by scale position (top-box teal, neutral gray, bottom-box red)
    top_vals = set(top_box_spec.get("values", []) if top_box_spec else [])
    bot_vals = set(bottom_box_spec.get("values", []) if bottom_box_spec else [])
    top_idx = bot_idx = 0
    seg_colors: list[str] = []
    for seg in segments:
        if seg["label"] in top_vals:
            seg_colors.append(_TOP_COLORS[min(top_idx, len(_TOP_COLORS) - 1)])
            top_idx += 1
        elif seg["label"] in bot_vals:
            seg_colors.append(_BOT_COLORS[min(bot_idx, len(_BOT_COLORS) - 1)])
            bot_idx += 1
        else:
            seg_colors.append(theme.GRAY_MID)

    chart_buf = chart_styles.render_stacked_bar(row_labels, segments, colors=seg_colors)

    cl, ct, cw, ch = theme.CHART_BOX
    chart_buf.seek(0)
    slide.shapes.add_picture(chart_buf,
                             Inches(cl), Inches(ct),
                             width=Inches(cw), height=Inches(ch))

    # ── Legend below chart ────────────────────────────────────────────────────
    legend_labels = [seg["label"] for seg in segments]
    leg_buf = chart_styles.legend_image(legend_labels, colors=seg_colors)
    leg_buf.seek(0)
    leg_t = ct + ch + 0.05
    slide.shapes.add_picture(leg_buf,
                             Inches(cl), Inches(leg_t),
                             width=Inches(cw), height=Inches(0.28))

    # ── Annotation panel ─────────────────────────────────────────────────────
    if annot_lines:
        _add_multiline_textbox(slide, theme.ANNOT_BOX, annot_lines,
                                font_size=theme.FONT_BODY, color=theme.GRAY_DARK)

    _add_footer(slide, base_note, page_num)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_pptx(
    results: list[QuestionResult],
    annotations: dict,
    output_path: str,
    logo_path: str | None = None,
) -> str:
    prs = Presentation()
    prs.slide_width  = Inches(theme.SLIDE_W)
    prs.slide_height = Inches(theme.SLIDE_H)

    result_map = {r.variable: r for r in results}
    page_num   = 1

    for spec in annotations.get("slides", []):
        slide_type = spec.get("type", "chart")

        if slide_type == "cover":
            add_cover_slide(prs, spec, logo_path=logo_path)

        elif slide_type == "section":
            add_section_slide(prs, spec, logo_path=logo_path)

        elif slide_type == "commentary":
            add_commentary_slide(prs, spec, page_num=page_num)
            page_num += 1

        elif slide_type == "chart":
            result = result_map.get(spec.get("variable", ""))
            add_chart_slide(prs, spec, result,
                            page_num=page_num, logo_path=logo_path)
            page_num += 1

        elif slide_type == "grid":
            # Resolve variable list to QuestionResult objects
            grid_results = [
                result_map[v]
                for v in spec.get("variables", [])
                if v in result_map
            ]
            add_grid_slide(prs, spec, grid_results,
                           page_num=page_num, logo_path=logo_path)
            page_num += 1

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
