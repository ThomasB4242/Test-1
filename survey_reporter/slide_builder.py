"""Assemble a PowerPoint presentation from QuestionResult objects and YAML annotations."""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

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
    """Add a text box with multiple paragraphs (one per line)."""
    l, t, w, h = ltwh
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = _rgb(color)
        run.font.bold = bold
    return txBox


def _add_filled_rect(slide, ltwh, fill_color: str) -> Any:
    from pptx.util import Inches as I
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    l, t, w, h = ltwh
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        I(l), I(t), I(w), I(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_color)
    shape.line.fill.background()
    return shape


def _add_oval(slide, ltwh, fill_color: str) -> Any:
    from pptx.util import Inches as I
    l, t, w, h = ltwh
    shape = slide.shapes.add_shape(
        9,  # MSO_SHAPE_TYPE.OVAL
        I(l), I(t), I(w), I(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_color)
    shape.line.fill.background()
    return shape


def _add_image(slide, buf: io.BytesIO, ltwh) -> Any:
    l, t, w, h = ltwh
    buf.seek(0)
    pic = slide.shapes.add_picture(buf, Inches(l), Inches(t),
                                   width=Inches(w), height=Inches(h))
    return pic


def _blank_slide(prs: Presentation):
    blank_layout = prs.slide_layouts[6]  # completely blank
    return prs.slides.add_slide(blank_layout)


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

def _add_header_stripe(slide):
    """Dark teal stripe across the top (logo area)."""
    _add_filled_rect(slide, (0, 0, theme.SLIDE_W, 0.50), theme.TEAL_DARK)


def _add_footer(slide, base_note: str, page_num: int | None):
    """Footer line with base note and optional page number."""
    _add_textbox(slide, theme.FOOTER_BOX, base_note,
                 font_size=theme.FONT_FOOTER, color=theme.GRAY_DARK, italic=True)
    if page_num is not None:
        _add_textbox(slide, theme.PAGE_NUM_BOX, str(page_num),
                     font_size=theme.FONT_FOOTER, color=theme.WHITE,
                     align=PP_ALIGN.RIGHT)


def add_cover_slide(prs: Presentation, spec: dict, logo_path: str | None = None):
    slide = _blank_slide(prs)
    # Full-bleed teal background
    _add_filled_rect(slide, (0, 0, theme.SLIDE_W, theme.SLIDE_H), theme.TEAL_DARK)

    title = spec.get("title", "Survey Report")
    subtitle = spec.get("subtitle", "")
    base = spec.get("base", "")

    # Title
    _add_textbox(slide, (0.5, 2.5, 9.0, 1.4), title,
                 font_size=32, color=theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
    # Subtitle
    if subtitle:
        _add_textbox(slide, (0.5, 3.9, 9.0, 0.7), subtitle,
                     font_size=18, color=theme.TEAL_LIGHT, align=PP_ALIGN.CENTER)
    # Base / n
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

    if logo_path and Path(logo_path).exists():
        try:
            slide.shapes.add_picture(logo_path, Inches(0.25), Inches(6.8),
                                     height=Inches(0.35))
        except Exception:
            pass


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
    bullets = spec.get("bullets", [])
    for b in bullets:
        lines.append(f"• {b}" if not b.startswith("•") else b)

    if lines:
        _add_multiline_textbox(
            slide,
            (0.30, 1.30, 9.40, 5.80),
            lines,
            font_size=theme.FONT_BODY,
            color=theme.GRAY_DARK,
        )

    base_note = spec.get("base", "")
    _add_footer(slide, base_note, page_num)


def add_chart_slide(
    prs: Presentation,
    spec: dict,
    result: QuestionResult | None,
    page_num: int | None = None,
    logo_path: str | None = None,
):
    slide = _blank_slide(prs)
    _add_header_stripe(slide)

    heading = spec.get("heading", result.label if result else "")
    question = spec.get("question", "")
    chart_type = spec.get("chart_type", "bar")
    annotations_lines = spec.get("annotations", [])
    base_note = spec.get("base", "")

    # Logo in header
    if logo_path and Path(logo_path).exists():
        try:
            slide.shapes.add_picture(logo_path, Inches(theme.LOGO_BOX[0]),
                                     Inches(theme.LOGO_BOX[1]),
                                     height=Inches(theme.LOGO_BOX[3]))
        except Exception:
            pass

    # Heading
    _add_textbox(slide, theme.HEADING_BOX, heading,
                 font_size=theme.FONT_HEADING, color=theme.CORAL, bold=True)

    # Question text
    if question:
        _add_textbox(slide, theme.QUESTION_BOX, question,
                     font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    # --- Chart ---
    chart_buf = None
    if result is not None and result.type == "categorical" and result.frequencies:
        freqs = result.frequencies
        labels = [f.label for f in freqs]
        percents = [f.percent for f in freqs]

        if chart_type == "stacked":
            # Build one segment per response option (treat each as its own stacked row
            # across a single implied "all respondents" bar set) — but for a stacked
            # chart we need multiple rows (one per category label).
            # Here each frequency row becomes a category with one coloured segment.
            # We render it as a stacked bar across all responses (single stacked bar
            # per "all respondents", labelled by option).
            segments = [
                {"label": f.label, "values": [f.percent]}
                for f in freqs
            ]
            single_categories = [""]
            chart_buf = chart_styles.render_stacked_bar(
                single_categories, segments,
                colors=theme.LIKERT_COLORS[:len(segments)],
            )
        else:
            # Reverse so largest bar is at top
            labels_rev = list(reversed(labels))
            percents_rev = list(reversed(percents))
            chart_buf = chart_styles.render_simple_bar(labels_rev, percents_rev)

    if chart_buf is not None:
        cl, ct, cw, ch = theme.CHART_BOX
        chart_buf.seek(0)
        slide.shapes.add_picture(chart_buf,
                                 Inches(cl), Inches(ct),
                                 width=Inches(cw), height=Inches(ch))

    # --- Combined % circles ---
    top_box_spec = spec.get("top_box")
    bottom_box_spec = spec.get("bottom_box")

    if result is not None and result.type == "categorical":
        freq_by_label = {f.label: f.percent for f in result.frequencies}

        if top_box_spec:
            pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
            label = top_box_spec.get("label", "")
            _draw_circle_badge(slide, theme.TOP_CIRCLE, theme.TEAL_MID,
                               pct, label)

        if bottom_box_spec:
            pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
            label = bottom_box_spec.get("label", "")
            _draw_circle_badge(slide, theme.BOTTOM_CIRCLE, theme.CORAL,
                               pct, label)

    # --- Annotations panel ---
    if annotations_lines:
        _add_multiline_textbox(
            slide, theme.ANNOT_BOX,
            annotations_lines,
            font_size=theme.FONT_BODY,
            color=theme.GRAY_DARK,
        )

    _add_footer(slide, base_note, page_num)


def _sum_box(freq_by_label: dict, values: list[str]) -> int:
    total = sum(freq_by_label.get(v, 0.0) for v in values)
    return int(round(total))


def _draw_circle_badge(slide, ltwh, fill_color: str, pct: int, label: str):
    """Draw a filled circle with pct% and label text centred inside it."""
    l, t, w, h = ltwh
    _add_oval(slide, (l, t, w, h), fill_color)

    # Percentage number
    _add_textbox(slide, (l, t + h * 0.12, w, h * 0.55),
                 f"{pct}%",
                 font_size=theme.FONT_CIRCLE, color=theme.WHITE,
                 bold=True, align=PP_ALIGN.CENTER)
    # Label text below
    _add_textbox(slide, (l, t + h * 0.62, w, h * 0.32),
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
    """Assemble the full .pptx and save it to *output_path*.

    Parameters
    ----------
    results : list[QuestionResult]
        Analysis output from analyzer.analyze().
    annotations : dict
        Parsed YAML annotations (from annotations.load_annotations() or auto_annotations()).
    output_path : str
        Destination .pptx file path.
    logo_path : str | None
        Optional path to a logo image file.

    Returns
    -------
    str  The resolved output_path.
    """
    prs = Presentation()
    prs.slide_width = Inches(theme.SLIDE_W)
    prs.slide_height = Inches(theme.SLIDE_H)

    # Build a lookup from variable name → QuestionResult
    result_map = {r.variable: r for r in results}

    slides_spec = annotations.get("slides", [])
    page_num = 1

    for spec in slides_spec:
        slide_type = spec.get("type", "chart")

        if slide_type == "cover":
            add_cover_slide(prs, spec, logo_path=logo_path)

        elif slide_type == "section":
            add_section_slide(prs, spec, logo_path=logo_path)

        elif slide_type == "commentary":
            add_commentary_slide(prs, spec, page_num=page_num)
            page_num += 1

        elif slide_type == "chart":
            variable = spec.get("variable", "")
            result = result_map.get(variable)
            add_chart_slide(prs, spec, result,
                            page_num=page_num, logo_path=logo_path)
            page_num += 1

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
