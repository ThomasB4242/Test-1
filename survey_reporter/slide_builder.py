"""Assemble a PowerPoint presentation from QuestionResult objects and YAML annotations."""
from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
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

_TOP_COLORS = [theme.TEAL_DARK, theme.TEAL_MID, theme.TEAL_LIGHT]
_BOT_COLORS = [theme.SALMON, theme.CORAL]


def _scale_bar_colors(labels_t2b: list[str],
                      top_box_spec: dict | None,
                      bottom_box_spec: dict | None) -> list[str]:
    """Per-bar hex colours based on top/bottom box membership."""
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
# Dynamic chart sizing
# ---------------------------------------------------------------------------

def _chart_h_for_n(n: int) -> float:
    """Ideal chart figure height (inches) for n bars / grid rows."""
    return float(min(5.20, max(1.80, n * 0.75 + 0.90)))


# ---------------------------------------------------------------------------
# Circle / bar position helpers
# ---------------------------------------------------------------------------

def _bar_circle_center_y(label_index: float, n_bars: int,
                         chart_top: float, chart_h: float,
                         axis_bottom: float | None = None) -> float:
    """Slide y-coordinate (inches) for the centre of bar at *label_index*.

    With invert_yaxis, index 0 is at the TOP.
    """
    if axis_bottom is None:
        axis_bottom = chart_styles.AXIS_BOTTOM
    ax_top = chart_top + (1.0 - chart_styles.AXIS_TOP) * chart_h
    ax_h   = (chart_styles.AXIS_TOP - axis_bottom) * chart_h
    return ax_top + (label_index + 0.5) * (ax_h / n_bars)


def _circle_cx_for_pct(
    pct: float,
    chart_left: float,
    chart_w: float,
    axis_left: float = chart_styles.AXIS_LEFT,
) -> float:
    """Slide x-coordinate (inches) for a circle at *pct* along the x-axis."""
    ax_w_frac = chart_styles.AXIS_RIGHT - axis_left
    x_frac    = axis_left + (pct / 100.0) * ax_w_frac
    return chart_left + x_frac * chart_w


def _grid_circle_r(n_rows: int, chart_h: float) -> float:
    """Radius (inches) sized to 40% of per-row height, capped at 0.36"."""
    ax_h      = (chart_styles.AXIS_TOP - chart_styles.GRID_AXIS_BOTTOM) * chart_h
    per_row_h = ax_h / n_rows
    return float(min(0.36, per_row_h * 0.40))


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
    run.font.name = theme.FONT_FACE
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
        run.font.name = theme.FONT_FACE
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
    for layout in prs.slide_layouts:
        if layout.name.lower() == "blank":
            return prs.slides.add_slide(layout)
    return prs.slides.add_slide(prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)])


def _load_template(template_path: str) -> Presentation:
    """Load a .pptx/.potm/.potx file as a base template with no slides."""
    p = template_path.lower()
    if p.endswith(".potm") or p.endswith(".potx"):
        buf = io.BytesIO()
        with zipfile.ZipFile(template_path, "r") as zin:
            with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "[Content_Types].xml":
                        data = data.replace(
                            b"application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
                            b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                        ).replace(
                            b"application/vnd.ms-powerpoint.template.main+xml",
                            b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                        )
                    zout.writestr(item, data)
        buf.seek(0)
        prs = Presentation(buf)
    else:
        prs = Presentation(template_path)

    # Remove all existing slides, keeping masters and layouts.
    # Access prs.slides first to trigger internal initialisation (rename_slide_parts),
    # then clear the slide-id list and drop the now-stale relationships.
    slides = prs.slides
    sldIdLst = slides._sldIdLst
    slide_rtype = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    )
    rids = [
        rId
        for rId, rel in prs.part.rels.items()
        if rel.reltype == slide_rtype
    ]
    while len(sldIdLst):
        sldIdLst.remove(sldIdLst[0])
    for rId in rids:
        prs.part._rels._rels.pop(rId, None)

    return prs


# ---------------------------------------------------------------------------
# Shared slide furniture
# ---------------------------------------------------------------------------

def _add_heading_rule(slide):
    """Thin teal rule under the heading area — replaces the old full-width stripe."""
    _add_filled_rect(slide, (0.34, 1.10, theme.SLIDE_W - 0.34, 0.03), theme.TEAL_DARK)


def _add_footer(slide, base_note: str, page_num: int | None):
    # Footer sits above the master logo (logo is at y≈6.85", h≈0.40")
    _add_textbox(slide, theme.FOOTER_BOX, base_note,
                 font_size=theme.FONT_FOOTER, color=theme.GRAY_DARK, italic=True)
    if page_num is not None:
        _add_textbox(slide, theme.PAGE_NUM_BOX, str(page_num),
                     font_size=theme.FONT_FOOTER, color=theme.GRAY_DARK,
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
                       fill_color: str, pct: int,
                       label: str = "", show_label: bool = True):
    """Draw a filled circle centred at (cx, cy) with radius r (inches).

    When show_label=False the percentage number is centred in the circle
    and no sub-label is drawn (use a legend entry to explain the circle).
    """
    l, t, d = cx - r, cy - r, r * 2
    _add_oval(slide, (l, t, d, d), fill_color)

    # Scale font size with circle radius
    pct_fs    = max(9, int(r * 52))
    label_fs  = max(7, int(r * 28))

    if show_label and label:
        _add_textbox(slide, (l, t + d * 0.10, d, d * 0.52),
                     f"{pct}%",
                     font_size=pct_fs, color=theme.WHITE,
                     bold=True, align=PP_ALIGN.CENTER)
        _add_textbox(slide, (l, t + d * 0.58, d, d * 0.36),
                     label,
                     font_size=label_fs, color=theme.WHITE,
                     align=PP_ALIGN.CENTER)
    else:
        # Number only — vertically centred
        _add_textbox(slide, (l, t + d * 0.22, d, d * 0.56),
                     f"{pct}%",
                     font_size=pct_fs, color=theme.WHITE,
                     bold=True, align=PP_ALIGN.CENTER)


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
    _add_textbox(slide, (0.67, 2.5, 12.0, 1.4), title,
                 font_size=32, color=theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
    if subtitle:
        _add_textbox(slide, (0.67, 3.9, 12.0, 0.7), subtitle,
                     font_size=18, color=theme.TEAL_LIGHT, align=PP_ALIGN.CENTER)
    if base:
        _add_textbox(slide, (0.67, 6.8, 12.0, 0.4), base,
                     font_size=10, color=theme.TEAL_LIGHT, align=PP_ALIGN.CENTER)
    if logo_path and Path(logo_path).exists():
        try:
            slide.shapes.add_picture(logo_path, Inches(0.33), Inches(0.15),
                                     height=Inches(0.35))
        except Exception:
            pass


def add_section_slide(prs: Presentation, spec: dict, logo_path: str | None = None):
    slide = _blank_slide(prs)
    _add_filled_rect(slide, (0, 0, theme.SLIDE_W, theme.SLIDE_H), theme.TEAL_DARK)
    title = spec.get("title", "Section")
    _add_textbox(slide, (0.67, 3.0, 12.0, 1.2), title,
                 font_size=28, color=theme.WHITE, bold=True, align=PP_ALIGN.LEFT)


def add_commentary_slide(prs: Presentation, spec: dict, page_num: int | None = None):
    slide = _blank_slide(prs)
    _add_heading_rule(slide)
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
        _add_multiline_textbox(slide, (0.40, 1.30, 12.53, 5.80), lines,
                                font_size=theme.FONT_BODY, color=theme.GRAY_DARK)
    _add_footer(slide, spec.get("base", ""), page_num)


# ---------------------------------------------------------------------------
# Layout constants
# ---------------------------------------------------------------------------

_CHART_L  = theme.CHART_BOX[0]   # 0.34"
_CHART_W  = theme.CHART_BOX[2]   # 7.33"
_CHART_T  = 1.60
_ANNOT_L  = 8.13
_ANNOT_W  = 12.93 - _ANNOT_L     # ~4.80"

# Grid uses a narrower chart so circles stay within the slide footprint
_GRID_CHART_W  = 6.67
_GRID_ANNOT_L  = 7.73
_GRID_ANNOT_W  = 12.93 - _GRID_ANNOT_L  # ~5.20"


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
    _add_heading_rule(slide)
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

    # ── Sizing ───────────────────────────────────────────────────────────────
    cl, ct, cw = _CHART_L, _CHART_T, _CHART_W
    n_bars     = 0
    labels_t2b = []
    chart_buf  = None

    if result is not None and result.type == "categorical" and result.frequencies:
        freqs = result.frequencies

        if chart_type == "stacked":
            row_labels = [""]
            segments   = [{"label": f.label, "values": [f.percent]} for f in freqs]
            colors     = theme.LIKERT_COLORS[:len(segments)]
            ch         = _chart_h_for_n(1)
            chart_buf  = chart_styles.render_stacked_bar(
                row_labels, segments, colors=colors, fig_h=ch)
            n_bars = 1
            labels_t2b = [""]
        else:
            labels_t2b   = list(reversed([f.label   for f in freqs]))
            percents_t2b = list(reversed([f.percent for f in freqs]))
            n_bars       = len(labels_t2b)
            ch           = _chart_h_for_n(n_bars)
            bar_colors   = _scale_bar_colors(labels_t2b, top_box_spec, bottom_box_spec)
            chart_buf    = chart_styles.render_simple_bar(
                labels_t2b, percents_t2b, bar_colors=bar_colors, fig_h=ch)
    else:
        ch = _chart_h_for_n(5)

    if chart_buf is not None:
        chart_buf.seek(0)
        slide.shapes.add_picture(chart_buf,
                                 Inches(cl), Inches(ct),
                                 width=Inches(cw), height=Inches(ch))

    # ── Circle badges ────────────────────────────────────────────────────────
    if result is not None and result.type == "categorical":
        freq_by_label = {f.label: f.percent for f in result.frequencies}
        r  = 0.62
        cx = cl + cw - r * 0.4  # overlaps right edge of chart

        if chart_type == "bar" and n_bars > 0:
            if top_box_spec:
                idxs = [i for i, lbl in enumerate(labels_t2b)
                        if lbl in set(top_box_spec.get("values", []))]
                if idxs:
                    cy  = _bar_circle_center_y(
                        sum(idxs) / len(idxs), n_bars, ct, ch)
                    pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
                    _draw_circle_badge(slide, cx, cy, r, theme.TEAL_MID,
                                       pct, top_box_spec.get("label", ""))

            if bottom_box_spec:
                idxs = [i for i, lbl in enumerate(labels_t2b)
                        if lbl in set(bottom_box_spec.get("values", []))]
                if idxs:
                    cy  = _bar_circle_center_y(
                        sum(idxs) / len(idxs), n_bars, ct, ch)
                    pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
                    _draw_circle_badge(slide, cx, cy, r * 0.85, theme.CORAL,
                                       pct, bottom_box_spec.get("label", ""))

    # ── Annotation panel ─────────────────────────────────────────────────────
    if annot_lines:
        _add_multiline_textbox(
            slide, (_ANNOT_L, ct, _ANNOT_W, ch),
            annot_lines, font_size=theme.FONT_BODY, color=theme.GRAY_DARK)

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
    """Stacked horizontal bar chart — one row per question.

    Circle badges are positioned at the actual pct% mark on each row's bar.
    Only the number is shown in the circle; the legend explains the colour.
    """
    slide = _blank_slide(prs)
    _add_heading_rule(slide)
    _add_logo(slide, logo_path)

    heading             = spec.get("heading", "")
    question            = spec.get("question", "")
    annot_lines         = spec.get("annotations", [])
    base_note           = spec.get("base", "")
    top_box_spec        = spec.get("top_box")
    bottom_box_spec     = spec.get("bottom_box")
    scale_order         = spec.get("scale_order", [])
    row_label_overrides = spec.get("row_labels", [])

    _add_textbox(slide, theme.HEADING_BOX, heading,
                 font_size=theme.FONT_HEADING, color=theme.CORAL, bold=True)
    if question:
        _add_textbox(slide, theme.QUESTION_BOX, question,
                     font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    if not results:
        _add_footer(slide, base_note, page_num)
        return

    n_rows = len(results)
    cl, ct = _CHART_L, _CHART_T
    cw     = _GRID_CHART_W
    ch     = _chart_h_for_n(n_rows)

    if not scale_order:
        scale_order = [f.label for f in results[0].frequencies]

    # Row labels (override or SPSS label, truncated if needed)
    row_labels = []
    for i, result in enumerate(results):
        if i < len(row_label_overrides) and row_label_overrides[i]:
            row_labels.append(row_label_overrides[i])
        else:
            lbl = result.label
            row_labels.append(lbl[:50] + "…" if len(lbl) > 50 else lbl)

    # Build segments
    segments: list[dict] = []
    for opt_label in scale_order:
        values = []
        for result in results:
            pct = next((f.percent for f in result.frequencies
                        if f.label == opt_label), 0.0)
            values.append(pct)
        segments.append({"label": opt_label, "values": values})

    # Segment colours
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

    # ── Render chart ─────────────────────────────────────────────────────────
    chart_buf = chart_styles.render_stacked_bar(
        row_labels, segments, colors=seg_colors,
        fig_h=ch,
        axis_left=chart_styles.GRID_AXIS_LEFT,
        axis_bottom=chart_styles.GRID_AXIS_BOTTOM,
    )
    chart_buf.seek(0)
    slide.shapes.add_picture(chart_buf,
                             Inches(cl), Inches(ct),
                             width=Inches(cw), height=Inches(ch))

    # ── Per-row circle badges at actual pct% along bar ────────────────────────
    r             = _grid_circle_r(n_rows, ch)
    top_vals_list = top_box_spec.get("values", []) if top_box_spec else []

    for row_i, result in enumerate(results):
        if not top_vals_list:
            continue
        freq_by_label = {f.label: f.percent for f in result.frequencies}
        pct = _sum_box(freq_by_label, top_vals_list)
        if pct <= 0:
            continue
        cx = _circle_cx_for_pct(pct, cl, cw, axis_left=chart_styles.GRID_AXIS_LEFT)
        cy = _bar_circle_center_y(row_i, n_rows, ct, ch,
                                  axis_bottom=chart_styles.GRID_AXIS_BOTTOM)
        _draw_circle_badge(slide, cx, cy, r, theme.TEAL_MID, pct,
                           show_label=False)

    # ── Legend (circle for total, rectangles for segments) ───────────────────
    circle_label  = top_box_spec.get("label", "Total") if top_box_spec else None
    legend_labels = [seg["label"] for seg in segments]
    leg_colors    = list(seg_colors)
    circle_idx: set[int] = set()
    if circle_label:
        legend_labels = [circle_label] + legend_labels
        leg_colors    = [theme.TEAL_MID] + leg_colors
        circle_idx    = {0}

    leg_h   = 0.30
    leg_t   = ct + ch + 0.08
    leg_buf = chart_styles.legend_image(
        legend_labels, colors=leg_colors,
        fig_w=cw, circle_indices=circle_idx,
    )
    leg_buf.seek(0)
    slide.shapes.add_picture(leg_buf,
                             Inches(cl), Inches(leg_t),
                             width=Inches(cw), height=Inches(leg_h))

    # ── Annotation panel ──────────────────────────────────────────────────────
    if annot_lines:
        _add_multiline_textbox(
            slide, (_GRID_ANNOT_L, ct, _GRID_ANNOT_W, ch),
            annot_lines, font_size=theme.FONT_BODY, color=theme.GRAY_DARK)

    _add_footer(slide, base_note, page_num)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_pptx(
    results: list[QuestionResult],
    annotations: dict,
    output_path: str,
    logo_path: str | None = None,
    template_path: str | None = None,
) -> str:
    if template_path and Path(template_path).exists():
        prs = _load_template(template_path)
    else:
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
