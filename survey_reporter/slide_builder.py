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
# Circle positioning helper
# ---------------------------------------------------------------------------

def _bar_circle_center_y(
    label_index: float,
    n_bars: int,
    chart_top: float,
    chart_h: float,
) -> float:
    """Return slide y-coordinate (inches) for the center of bar at *label_index*.

    Uses the known subplots_adjust margins from chart_styles so that the
    circle lines up with the actual rendered bar position.

    *label_index* may be fractional (average of multiple bar indices).
    With invert_yaxis, index 0 is displayed at the TOP of the chart.
    """
    # Axis occupies AXIS_TOP..AXIS_BOTTOM fractions of the figure (from bottom).
    # In slide coords the axis top is (1 - AXIS_TOP) * chart_h from chart_top.
    ax_top_slide = chart_top + (1.0 - chart_styles.AXIS_TOP) * chart_h
    ax_h_slide   = (chart_styles.AXIS_TOP - chart_styles.AXIS_BOTTOM) * chart_h
    bar_slot     = ax_h_slide / n_bars
    # index 0 is at top (invert_yaxis), so y increases downward
    return ax_top_slide + (label_index + 0.5) * bar_slot


def _circle_x_center(chart_left: float, chart_w: float, radius: float) -> float:
    """Center the circle so its right edge lines up with the chart right edge."""
    return chart_left + chart_w - radius * 0.4


# ---------------------------------------------------------------------------
# Text-box helpers
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
# Chart slide
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

    heading     = spec.get("heading", result.label if result else "")
    question    = spec.get("question", "")
    chart_type  = spec.get("chart_type", "bar")
    annot_lines = spec.get("annotations", [])
    base_note   = spec.get("base", "")
    top_box_spec    = spec.get("top_box")
    bottom_box_spec = spec.get("bottom_box")

    # Logo
    if logo_path and Path(logo_path).exists():
        try:
            slide.shapes.add_picture(logo_path,
                                     Inches(theme.LOGO_BOX[0]), Inches(theme.LOGO_BOX[1]),
                                     height=Inches(theme.LOGO_BOX[3]))
        except Exception:
            pass

    # Heading + question text
    _add_textbox(slide, theme.HEADING_BOX, heading,
                 font_size=theme.FONT_HEADING, color=theme.CORAL, bold=True)
    if question:
        _add_textbox(slide, theme.QUESTION_BOX, question,
                     font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    # ── Build chart ──────────────────────────────────────────────────────────
    chart_buf    = None
    n_bars       = 0
    labels_t2b   = []   # labels top-to-bottom (for circle positioning)

    if result is not None and result.type == "categorical" and result.frequencies:
        freqs = result.frequencies

        if chart_type == "stacked":
            # Each frequency row is one response option that becomes a colour segment.
            # For a true GRID question, the caller supplies multiple variables
            # (future feature). For now, show all options as a single stacked bar
            # with one row labeled "".
            row_labels = [""]
            segments = [
                {"label": f.label, "values": [f.percent]}
                for f in freqs
            ]
            chart_buf = chart_styles.render_stacked_bar(
                row_labels, segments,
                colors=theme.LIKERT_COLORS[:len(segments)],
            )
            n_bars = 1
            labels_t2b = [""]

        else:  # "bar" — individual bar per response option
            # Display order: reversed SPSS code order puts the most-positive
            # response (highest code value) at the top of the chart.
            labels_t2b   = list(reversed([f.label   for f in freqs]))
            percents_t2b = list(reversed([f.percent for f in freqs]))
            n_bars = len(labels_t2b)
            chart_buf = chart_styles.render_simple_bar(labels_t2b, percents_t2b)

    # Embed chart image at EXACT CHART_BOX dimensions (chart was rendered to match)
    cl, ct, cw, ch = theme.CHART_BOX
    if chart_buf is not None:
        chart_buf.seek(0)
        slide.shapes.add_picture(chart_buf,
                                 Inches(cl), Inches(ct),
                                 width=Inches(cw), height=Inches(ch))

    # ── Combined-% circle badges ─────────────────────────────────────────────
    if result is not None and result.type == "categorical":
        freq_by_label = {f.label: f.percent for f in result.frequencies}

        circle_r = 0.72   # radius in inches
        # x: center is just inside the chart right edge so it overlaps the border
        cx = _circle_x_center(cl, cw, circle_r)

        if chart_type == "bar" and n_bars > 0:
            # Compute y from actual bar positions
            if top_box_spec:
                top_vals = set(top_box_spec.get("values", []))
                idxs = [i for i, l in enumerate(labels_t2b) if l in top_vals]
                if idxs:
                    avg_idx = sum(idxs) / len(idxs)
                    cy = _bar_circle_center_y(avg_idx, n_bars, ct, ch)
                    pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
                    _draw_circle_badge(slide, cx, cy, circle_r,
                                       theme.TEAL_MID, pct, top_box_spec.get("label", ""))

            if bottom_box_spec:
                bot_vals = set(bottom_box_spec.get("values", []))
                idxs = [i for i, l in enumerate(labels_t2b) if l in bot_vals]
                if idxs:
                    avg_idx = sum(idxs) / len(idxs)
                    cy = _bar_circle_center_y(avg_idx, n_bars, ct, ch)
                    pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
                    _draw_circle_badge(slide, cx, cy, circle_r * 0.9,
                                       theme.CORAL, pct, bottom_box_spec.get("label", ""))

        else:
            # Stacked bar: use theme-defined positions
            if top_box_spec:
                pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
                _draw_circle_badge(slide, *theme.TOP_CIRCLE_CENTER_R,
                                   theme.TEAL_MID, pct, top_box_spec.get("label", ""))
            if bottom_box_spec:
                pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
                _draw_circle_badge(slide, *theme.BOTTOM_CIRCLE_CENTER_R,
                                   theme.CORAL, pct, bottom_box_spec.get("label", ""))

    # ── Annotations panel ────────────────────────────────────────────────────
    if annot_lines:
        _add_multiline_textbox(slide, theme.ANNOT_BOX, annot_lines,
                                font_size=theme.FONT_BODY, color=theme.GRAY_DARK)

    _add_footer(slide, base_note, page_num)


# ---------------------------------------------------------------------------
# Circle badge helper
# ---------------------------------------------------------------------------

def _sum_box(freq_by_label: dict, values: list[str]) -> int:
    return int(round(sum(freq_by_label.get(v, 0.0) for v in values)))


def _draw_circle_badge(slide, cx: float, cy: float, r: float,
                       fill_color: str, pct: int, label: str):
    """Draw a filled circle (cx, cy, r in inches) with pct% and label inside."""
    l = cx - r
    t = cy - r
    d = r * 2
    _add_oval(slide, (l, t, d, d), fill_color)

    # Percentage number — upper portion of circle
    _add_textbox(slide, (l, t + d * 0.10, d, d * 0.52),
                 f"{pct}%",
                 font_size=theme.FONT_CIRCLE, color=theme.WHITE,
                 bold=True, align=PP_ALIGN.CENTER)
    # Label text — lower portion of circle
    _add_textbox(slide, (l, t + d * 0.58, d, d * 0.36),
                 label,
                 font_size=theme.FONT_CIRCLE_SM, color=theme.WHITE,
                 align=PP_ALIGN.CENTER)


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

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
