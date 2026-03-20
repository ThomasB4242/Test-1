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


def _add_multiline_textbox(slide, ltwh, lines: list, font_size: int,
                           color: str, bold: bool = False) -> Any:
    """Add a multi-line textbox.

    Each item in *lines* may be:
    - a plain ``str``  → rendered in *color*
    - a ``dict`` with keys ``text`` (str) and optionally ``color``
      ("teal" → TEAL_MID, "coral" → CORAL, else *color* default)
    """
    _COLOR_MAP = {"teal": theme.TEAL_MID, "coral": theme.CORAL}

    l, t, w, h = ltwh
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, dict):
            text      = item.get("text", "")
            line_color = _COLOR_MAP.get(item.get("color", ""), color)
            line_bold  = item.get("bold", bold)
        else:
            text      = str(item)
            line_color = color
            line_bold  = bold
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = text
        run.font.name = theme.FONT_FACE
        run.font.size = Pt(font_size)
        run.font.color.rgb = _rgb(line_color)
        run.font.bold = line_bold
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


def _find_layout(prs: Presentation, *names: str):
    """Return the first slide layout whose name matches one of *names* (case-insensitive).
    Falls back to the Blank layout if none match."""
    lower = {n.lower() for n in names}
    for layout in prs.slide_layouts:
        if layout.name.lower() in lower:
            return layout
    return _blank_slide.__wrapped__(prs) if hasattr(_blank_slide, '__wrapped__') else None


def _layout_slide(prs: Presentation, *layout_names: str):
    """Add a slide using the named layout (falls back to Blank)."""
    lower = {n.lower() for n in layout_names}
    for layout in prs.slide_layouts:
        if layout.name.lower() in lower:
            return prs.slides.add_slide(layout)
    return _blank_slide(prs)


def _fill_placeholder(slide, ph_idx: int, text: str):
    """Write *text* into the slide placeholder with the given idx.

    Uses the template's built-in formatting — no explicit overrides.
    Returns the placeholder shape, or None if not found.
    """
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == ph_idx:
            ph.text_frame.text = text
            return ph
    return None


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
    # Footer text only — slide numbering is handled by the template master.
    if base_note:
        _add_textbox(slide, theme.FOOTER_BOX, base_note,
                     font_size=theme.FONT_FOOTER, color=theme.GRAY_DARK, italic=True)


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

    The percentage number and % sign use different font sizes (% is smaller).
    Label font is sized so the text always fits; the label textbox is made
    wider than the circle so nothing is clipped (excess is invisible on the
    slide background).
    """
    l, t, d = cx - r, cy - r, r * 2
    _add_oval(slide, (l, t, d, d), fill_color)

    # Scale font sizes with circle radius; % sign is noticeably smaller
    num_fs   = max(10, int(r * 62))   # main number
    pct_fs   = max(7,  int(r * 40))   # % suffix — smaller

    # Auto-size the label to guarantee it fits inside the circle.
    # Use 80% of circle width as the effective text area and 0.70 pts-per-char
    # (conservative for bold fonts) so the calculation always leaves room.
    label_fs_base = max(7, int(r * 32))
    if label:
        circle_pt_w = d * 72
        chars_fit   = max(3, int(circle_pt_w * 0.80 / (label_fs_base * 0.70)))
        label_fs    = label_fs_base if len(label) <= chars_fit else max(
            6, int(label_fs_base * chars_fit / len(label))
        )
    else:
        label_fs = label_fs_base

    def _pct_textbox(top_frac, h_frac):
        """Add a textbox with number + % in two runs at different sizes."""
        txBox = slide.shapes.add_textbox(
            Inches(l), Inches(t + d * top_frac), Inches(d), Inches(d * h_frac)
        )
        tf = txBox.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        for txt, fs in ((str(pct), num_fs), ("%", pct_fs)):
            run = p.add_run()
            run.text = txt
            run.font.name = theme.FONT_FACE
            run.font.size = Pt(fs)
            run.font.color.rgb = _rgb(theme.WHITE)
            run.font.bold = True

    if show_label and label:
        _pct_textbox(0.06, 0.52)
        # Make the label textbox 40% wider than the circle (20% extra each side)
        # so text never gets clipped. White text beyond the circle boundary is
        # invisible against the slide background.
        extra = d * 0.20
        _add_textbox(slide, (l - extra, t + d * 0.56, d + extra * 2, d * 0.38),
                     label, font_size=label_fs, color=theme.WHITE,
                     align=PP_ALIGN.CENTER, word_wrap=False)
    else:
        # Number only — vertically centred
        _pct_textbox(0.18, 0.60)


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
# Layout constants — matched to the Blank.potm template positions
# ---------------------------------------------------------------------------

_CHART_L  = 0.34    # chart image left edge
_CHART_W  = 5.80    # single-bar chart image width
_CHART_T  = 1.87    # chart image top (below title + question)
# Annotation: use template's preset right-panel position (ANNOT_BOX) at 14pt
_ANNOT_L  = theme.ANNOT_BOX[0]   # 8.07"  — centre of the teal diagonal panel
_ANNOT_W  = theme.ANNOT_BOX[2]   # 4.87"
_ANNOT_FONT = 14                  # user-requested font size for annotations

# Grid layout — wider chart, annotations pushed further right
# Strip on layout '3_Title Slide' runs y=2.23" to y=6.58" (h=4.35")
_GRID_MINT_TOP     = 2.23   # top of the mint-coloured strip
_GRID_MINT_BOTTOM  = 6.58   # bottom of the mint-coloured strip
_GRID_CHART_W      = 9.00   # grid stacked-bar image width
_GRID_CHART_MAX_H  = 4.30   # cap height so chart+legend fits in strip
_GRID_ANNOT_L      = 9.50   # grid annotation left (just past chart)
_GRID_ANNOT_W      = 12.93 - _GRID_ANNOT_L  # ~3.43"

# Template layout names
_LAYOUT_SINGLE_BAR = "title slide"       # Layout 0  — diagonal right background
_LAYOUT_GRID       = "3_title slide"     # Layout 11 — full-width mint horizontal strip
_LAYOUT_TABLE      = "1_title slide"     # Layout 10 — white solid background, same placeholders


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
    # Use the 'Title Slide' layout — provides the diagonal-right background design
    slide = _layout_slide(prs, _LAYOUT_SINGLE_BAR)
    _add_logo(slide, logo_path)

    heading         = spec.get("heading", result.label if result else "")
    question        = spec.get("question", "")
    chart_type      = spec.get("chart_type", "bar")
    annot_lines     = spec.get("annotations", [])
    base_note       = spec.get("base", "")
    top_box_spec    = spec.get("top_box")
    bottom_box_spec = spec.get("bottom_box")

    # Fill template placeholders (idx 0=title, 21=question, 22=base/footer)
    if not _fill_placeholder(slide, 0, heading):
        _add_textbox(slide, theme.HEADING_BOX, heading,
                     font_size=theme.FONT_HEADING, color=theme.TEAL_DARK, bold=True)
        _add_heading_rule(slide)

    if question:
        if not _fill_placeholder(slide, 21, question):
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
            ch         = max(_chart_h_for_n(1), 2.00)
            chart_buf  = chart_styles.render_stacked_bar(
                row_labels, segments, colors=colors, fig_h=ch,
                axis_left=0.05)  # no row label — use full width
            n_bars = 1
            labels_t2b = [""]
        elif chart_type == "pie":
            labels_t2b   = [f.label   for f in freqs]
            percents_t2b = [f.percent for f in freqs]
            n_bars       = len(labels_t2b)
            ch           = 4.5   # fixed height for pie
            # Only colour-code slices when top/bottom boxes give distinct colours;
            # otherwise pass None so render_pie_chart uses its LIKERT_COLORS default.
            if top_box_spec or bottom_box_spec:
                pie_colors = _scale_bar_colors(labels_t2b, top_box_spec, bottom_box_spec)
            else:
                pie_colors = None
            chart_buf    = chart_styles.render_pie_chart(
                labels_t2b, percents_t2b, colors=pie_colors, fig_h=ch, fig_w=cw)
        else:
            labels_t2b   = list(reversed([f.label   for f in freqs]))
            percents_t2b = list(reversed([f.percent for f in freqs]))
            n_bars       = len(labels_t2b)
            ch           = _chart_h_for_n(n_bars)
            bar_colors   = _scale_bar_colors(labels_t2b, top_box_spec, bottom_box_spec)
            chart_buf    = chart_styles.render_simple_bar(
                labels_t2b, percents_t2b, bar_colors=bar_colors, fig_h=ch, fig_w=cw)
    else:
        ch = _chart_h_for_n(5)

    if chart_buf is not None:
        chart_buf.seek(0)
        # fig generated at exactly cw × ch — specify both to avoid stretching
        slide.shapes.add_picture(chart_buf,
                                 Inches(cl), Inches(ct),
                                 width=Inches(cw), height=Inches(ch))

    # ── Circle badges (bar charts only — pie has numbers embedded) ───────────
    if result is not None and result.type == "categorical" and chart_type == "bar" and n_bars > 0:
        freq_by_label = {f.label: f.percent for f in result.frequencies}

        # Bar height in slide inches — size circle to be clearly readable
        ax_h = (chart_styles.AXIS_TOP - chart_styles.AXIS_BOTTOM) * ch
        bar_h_slide = 0.60 * ax_h / max(n_bars, 1)
        r = max(0.38, min(0.72, bar_h_slide * 1.00))

        # x-axis scale from render_simple_bar
        xlim_max = max(percents_t2b) * 1.12 + 3 if percents_t2b else 100.0

        def _circle_cx_simple(pct_val: float, offset: float = 0.0) -> float:
            """Map a data value to slide x-coordinate, with optional inch offset."""
            x_frac = chart_styles.AXIS_LEFT + (pct_val / xlim_max) * (
                chart_styles.AXIS_RIGHT - chart_styles.AXIS_LEFT
            )
            return cl + x_frac * cw + offset

        if top_box_spec:
            top_vals_set = set(top_box_spec.get("values", []))
            idxs = [i for i, lbl in enumerate(labels_t2b) if lbl in top_vals_set]
            if idxs:
                cy  = _bar_circle_center_y(
                    sum(idxs) / len(idxs), n_bars, ct, ch)
                pct = _sum_box(freq_by_label, top_box_spec.get("values", []))
                # x = just beyond end of the longest bar in the set
                max_bar = max((freq_by_label.get(lbl, 0) for lbl in top_vals_set), default=0)
                cx = _circle_cx_simple(max_bar, offset=r * 1.15)
                _draw_circle_badge(slide, cx, cy, r, theme.TEAL_MID,
                                   pct, top_box_spec.get("label", ""))

        if bottom_box_spec:
            bot_vals_set = set(bottom_box_spec.get("values", []))
            idxs = [i for i, lbl in enumerate(labels_t2b) if lbl in bot_vals_set]
            if idxs:
                cy  = _bar_circle_center_y(
                    sum(idxs) / len(idxs), n_bars, ct, ch)
                pct = _sum_box(freq_by_label, bottom_box_spec.get("values", []))
                max_bar = max((freq_by_label.get(lbl, 0) for lbl in bot_vals_set), default=0)
                cx = _circle_cx_simple(max_bar, offset=r * 0.90 * 1.15)
                _draw_circle_badge(slide, cx, cy, r * 0.90, theme.CORAL,
                                   pct, bottom_box_spec.get("label", ""))

    # ── Annotation panel — positioned in the right teal panel ────────────────
    if annot_lines:
        annot_font = spec.get("annotation_font", _ANNOT_FONT)
        _add_multiline_textbox(
            slide, (_ANNOT_L, ct, _ANNOT_W, ch),
            annot_lines, font_size=annot_font, color=theme.GRAY_DARK)

    # Base note: use template placeholder idx=22 if available, else manual footer
    if not _fill_placeholder(slide, 22, base_note):
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

    Uses the 'Section Header' template layout (coloured horizontal strip).
    Total circles are rendered inside the matplotlib figure.
    Rows are sorted descending by top-box total percentage.
    """
    # Use the '3_Title Slide' layout — full-width mint horizontal strip
    slide = _layout_slide(prs, _LAYOUT_GRID)
    _add_logo(slide, logo_path)

    heading             = spec.get("heading", "")
    question            = spec.get("question", "")
    annot_lines         = spec.get("annotations", [])
    base_note           = spec.get("base", "")
    top_box_spec        = spec.get("top_box")
    bottom_box_spec     = spec.get("bottom_box")
    scale_order         = spec.get("scale_order", [])
    row_label_overrides = spec.get("row_labels", [])

    # Fill template placeholders; fall back to manual textboxes if not found
    if not _fill_placeholder(slide, 0, heading):
        _add_textbox(slide, theme.HEADING_BOX, heading,
                     font_size=theme.FONT_HEADING, color=theme.TEAL_DARK, bold=True)
        _add_heading_rule(slide)
    if question:
        if not _fill_placeholder(slide, 21, question):
            _add_textbox(slide, theme.QUESTION_BOX, question,
                         font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    if not results:
        _add_footer(slide, base_note, page_num)
        return

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

    # Compute top-box totals and sort rows descending
    top_vals_list = top_box_spec.get("values", []) if top_box_spec else []
    total_percents: list[float] = []
    for res in results:
        freq_by_label = {f.label: f.percent for f in res.frequencies}
        tp = float(_sum_box(freq_by_label, top_vals_list)) if top_vals_list else 0.0
        total_percents.append(tp)

    if top_vals_list:
        indexed = sorted(
            enumerate(zip(total_percents, row_labels)),
            key=lambda x: (-x[1][0], x[0]),   # desc pct, stable original order
        )
        order          = [i for i, _ in indexed]
        total_percents = [total_percents[i] for i in order]
        results        = [results[i]        for i in order]
        row_labels     = [row_labels[i]     for i in order]

    n_rows = len(results)
    cl     = _CHART_L
    cw     = _GRID_CHART_W
    # All grids use at least 4.00" so small grids fill the mint strip
    ch     = min(_GRID_CHART_MAX_H, max(4.00, _chart_h_for_n(n_rows)))
    # Vertically centre the chart within the mint-coloured strip
    ct     = (_GRID_MINT_TOP + _GRID_MINT_BOTTOM - ch) / 2

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

    # ── Build legend spec — legend embedded inside the chart figure ───────────
    circle_label = top_box_spec.get("label", "Total agree") if top_box_spec else None
    legend_spec: list[tuple] = []
    if circle_label:
        legend_spec.append((circle_label, theme.TEAL_MID, True))
    for seg, clr in zip(segments, seg_colors):
        legend_spec.append((seg["label"], clr, False))

    chart_buf = chart_styles.render_stacked_bar(
        row_labels, segments, colors=seg_colors,
        fig_h=ch, fig_w=cw,
        axis_left=None,   # auto-computed per label length inside render_stacked_bar
        axis_bottom=chart_styles.GRID_AXIS_BOTTOM,
        total_percents=total_percents if top_vals_list else None,
        circle_color=theme.TEAL_MID,
        legend_spec=legend_spec if legend_spec else None,
    )
    chart_buf.seek(0)
    slide.shapes.add_picture(chart_buf,
                             Inches(cl), Inches(ct),
                             width=Inches(cw), height=Inches(ch))

    # ── Annotation panel ──────────────────────────────────────────────────────
    if annot_lines:
        annot_font = spec.get("annotation_font", _ANNOT_FONT)
        _add_multiline_textbox(
            slide, (_GRID_ANNOT_L, ct, _GRID_ANNOT_W, ch),
            annot_lines, font_size=annot_font, color=theme.GRAY_DARK)

    # Base: layout '3_Title Slide' uses idx=22 for footer (at y=7.08")
    if not _fill_placeholder(slide, 22, base_note):
        _add_footer(slide, base_note, page_num)


# ---------------------------------------------------------------------------
# Table slide  (coded open-end / ranked list — label + %)
# ---------------------------------------------------------------------------

def add_table_slide(
    prs: Presentation,
    spec: dict,
    result: QuestionResult | None,
    page_num: int | None = None,
    logo_path: str | None = None,
):
    """Two-column table slide: full response label (bold title + normal explanation) | %.

    Uses the '1_Title Slide' layout (white solid background, same title /
    question / footer placeholders as chart slides — no diagonal teal panel).
    Each label is split on ': ' so the short title is bold and the explanation
    is regular weight.  Rows are shown in descending % order.
    """
    from lxml import etree

    # '1_Title Slide' (layout 10): white bg, same ph positions as Title Slide
    slide = _layout_slide(prs, _LAYOUT_TABLE)
    _add_logo(slide, logo_path)

    heading   = spec.get("heading", result.label if result else "")
    question  = spec.get("question", "")
    base_note = spec.get("base", "")

    # Fill template placeholders — identical pattern to add_chart_slide
    if not _fill_placeholder(slide, 0, heading):
        _add_textbox(slide, theme.HEADING_BOX, heading,
                     font_size=theme.FONT_HEADING, color=theme.TEAL_DARK, bold=True)
        _add_heading_rule(slide)
    if question:
        if not _fill_placeholder(slide, 21, question):
            _add_textbox(slide, theme.QUESTION_BOX, question,
                         font_size=theme.FONT_QUESTION, color=theme.GRAY_DARK, italic=True)

    _NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    def _apply_bg(cell, hex_color: str):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for child in list(tcPr):
            tag = child.tag.split('}')[-1]
            if tag in ('solidFill', 'noFill'):
                tcPr.remove(child)
        sf  = etree.SubElement(tcPr, f'{{{_NS}}}solidFill')
        clr = etree.SubElement(sf,   f'{{{_NS}}}srgbClr')
        clr.set('val', hex_color.lstrip('#'))

    def _header_cell(cell, text: str, align=PP_ALIGN.LEFT):
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = False
        p  = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name  = theme.FONT_FACE
        run.font.size  = Pt(11)
        run.font.bold  = True
        run.font.color.rgb = _rgb(theme.WHITE)
        _apply_bg(cell, theme.TEAL_DARK)

    def _label_cell(cell, label: str, fg: str, bg: str):
        """Bold the title (before first ': '); leave the explanation normal."""
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = True
        p  = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        if ': ' in label:
            title, explanation = label.split(': ', 1)
            for txt, bold in ((title, True), (': ' + explanation, False)):
                run = p.add_run()
                run.text = txt
                run.font.name  = theme.FONT_FACE
                run.font.size  = Pt(10)
                run.font.bold  = bold
                run.font.color.rgb = _rgb(fg)
        else:
            run = p.add_run()
            run.text = label
            run.font.name  = theme.FONT_FACE
            run.font.size  = Pt(10)
            run.font.bold  = True
            run.font.color.rgb = _rgb(fg)
        _apply_bg(cell, bg)

    def _pct_cell(cell, pct: float, fg: str, bg: str):
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = False
        p  = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = f"{int(round(pct))}%"
        run.font.name  = theme.FONT_FACE
        run.font.size  = Pt(10)
        run.font.bold  = True
        run.font.color.rgb = _rgb(fg)
        _apply_bg(cell, bg)

    if result is not None and result.frequencies:
        freqs    = list(reversed(result.frequencies))   # descending %
        n_rows   = len(freqs)
        max_pct  = max(f.percent for f in freqs)

        tbl_l    = _CHART_L
        tbl_t    = _CHART_T
        tbl_w    = theme.SLIDE_W - tbl_l - 0.40   # full width, small right margin
        avail_h  = theme.SLIDE_H - tbl_t - 0.55
        row_h_in = min(0.42, avail_h / max(n_rows + 1, 1))
        tbl_h    = row_h_in * (n_rows + 1)

        tbl_frame = slide.shapes.add_table(
            n_rows + 1, 2,
            Inches(tbl_l), Inches(tbl_t),
            Inches(tbl_w), Inches(tbl_h),
        )
        tbl = tbl_frame.table
        tbl.columns[0].width = Inches(tbl_w * 0.85)
        tbl.columns[1].width = Inches(tbl_w * 0.15)

        _header_cell(tbl.cell(0, 0), "Response")
        _header_cell(tbl.cell(0, 1), "%", align=PP_ALIGN.CENTER)

        _ROW_ALT = "#E6F8F9"
        for i, freq in enumerate(freqs):
            bg     = theme.WHITE if i % 2 == 0 else _ROW_ALT
            fg_lbl = theme.TEAL_DARK if i == 0 else theme.GRAY_DARK
            fg_pct = theme.TEAL_DARK if freq.percent >= max_pct * 0.7 else theme.GRAY_DARK
            _label_cell(tbl.cell(i + 1, 0), freq.label, fg_lbl, bg)
            _pct_cell(tbl.cell(i + 1, 1), freq.percent, fg_pct, bg)

    if not _fill_placeholder(slide, 22, base_note):
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

        elif slide_type == "table":
            result = result_map.get(spec.get("variable", ""))
            add_table_slide(prs, spec, result,
                            page_num=page_num, logo_path=logo_path)
            page_num += 1

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
