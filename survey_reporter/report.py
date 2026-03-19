"""Assemble a PDF survey report using ReportLab Platypus."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Sequence

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from .analyzer import QuestionResult
from .visualizer import chart_for_question

# ── Colours ───────────────────────────────────────────────────────────────────
BRAND_BLUE = colors.HexColor("#2E75B6")
LIGHT_GREY = colors.HexColor("#F2F2F2")
MID_GREY = colors.HexColor("#BFBFBF")
TEXT_DARK = colors.HexColor("#1A1A1A")


# ── Styles ────────────────────────────────────────────────────────────────────
def _build_styles():
    base = getSampleStyleSheet()
    styles = {}

    styles["cover_title"] = ParagraphStyle(
        "cover_title",
        parent=base["Title"],
        fontSize=28,
        textColor=colors.white,
        alignment=TA_CENTER,
        spaceAfter=8,
        leading=34,
    )
    styles["cover_sub"] = ParagraphStyle(
        "cover_sub",
        parent=base["Normal"],
        fontSize=13,
        textColor=colors.HexColor("#D9E8F5"),
        alignment=TA_CENTER,
    )
    styles["h1"] = ParagraphStyle(
        "h1",
        parent=base["Heading1"],
        fontSize=14,
        textColor=BRAND_BLUE,
        spaceBefore=14,
        spaceAfter=4,
        leading=18,
    )
    styles["body"] = ParagraphStyle(
        "body",
        parent=base["Normal"],
        fontSize=9,
        textColor=TEXT_DARK,
        spaceAfter=3,
        leading=13,
    )
    styles["caption"] = ParagraphStyle(
        "caption",
        parent=base["Normal"],
        fontSize=8,
        textColor=colors.HexColor("#666666"),
        alignment=TA_CENTER,
        spaceAfter=6,
    )
    styles["stat_label"] = ParagraphStyle(
        "stat_label",
        parent=base["Normal"],
        fontSize=9,
        textColor=TEXT_DARK,
        alignment=TA_RIGHT,
    )
    styles["stat_value"] = ParagraphStyle(
        "stat_value",
        parent=base["Normal"],
        fontSize=9,
        textColor=BRAND_BLUE,
        alignment=TA_LEFT,
    )
    return styles


# ── Page templates ─────────────────────────────────────────────────────────────
def _cover_bg(canvas, doc):
    """Paint a solid blue cover page background."""
    canvas.saveState()
    w, h = A4
    canvas.setFillColor(BRAND_BLUE)
    canvas.rect(0, 0, w, h, fill=True, stroke=False)
    canvas.restoreState()


def _page_header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4
    # Header bar
    canvas.setFillColor(BRAND_BLUE)
    canvas.rect(0, h - 18 * mm, w, 18 * mm, fill=True, stroke=False)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(colors.white)
    canvas.drawString(1.5 * cm, h - 11 * mm, doc.report_title)
    # Footer
    canvas.setFillColor(MID_GREY)
    canvas.rect(0, 0, w, 10 * mm, fill=True, stroke=False)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#555555"))
    canvas.drawRightString(w - 1.5 * cm, 3 * mm, f"Page {doc.page}")
    canvas.drawString(1.5 * cm, 3 * mm, f"Generated {date.today().strftime('%d %B %Y')}")
    canvas.restoreState()


# ── Public API ─────────────────────────────────────────────────────────────────
def build_report(
    results: Sequence[QuestionResult],
    output_path: str | Path,
    title: str = "Survey Report",
    n_responses: int = 0,
) -> Path:
    """Generate a PDF at *output_path* from the list of QuestionResult objects."""
    output_path = Path(output_path)
    styles = _build_styles()

    # ── Document setup ────────────────────────────────────────────────────────
    doc = BaseDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=2.5 * cm,
        bottomMargin=1.8 * cm,
    )
    doc.report_title = title  # used by header/footer callback

    W, H = A4
    content_w = W - 3.6 * cm

    cover_frame = Frame(0, 0, W, H, id="cover")
    body_frame = Frame(
        1.8 * cm, 1.5 * cm,
        content_w, H - 4.0 * cm,
        id="body",
    )

    doc.addPageTemplates([
        PageTemplate(id="Cover", frames=[cover_frame], onPage=_cover_bg),
        PageTemplate(id="Body", frames=[body_frame], onPage=_page_header_footer),
    ])

    story: list = []

    # ── Cover page ────────────────────────────────────────────────────────────
    story.append(NextPageTemplate("Cover"))
    story.append(Spacer(1, 8 * cm))
    story.append(Paragraph(title, styles["cover_title"]))
    story.append(Spacer(1, 6 * mm))
    today_str = date.today().strftime("%d %B %Y")
    story.append(Paragraph(today_str, styles["cover_sub"]))
    if n_responses:
        story.append(Spacer(1, 4 * mm))
        story.append(Paragraph(f"N = {n_responses:,} responses", styles["cover_sub"]))
    story.append(PageBreak())

    # ── Summary page ──────────────────────────────────────────────────────────
    story.append(NextPageTemplate("Body"))
    story.append(Paragraph("Executive Summary", styles["h1"]))
    story.append(Spacer(1, 3 * mm))

    n_cat = sum(1 for r in results if r.type == "categorical")
    n_num = sum(1 for r in results if r.type == "numeric")

    summary_data = [
        ["Total responses", f"{n_responses:,}"],
        ["Questions analysed", str(len(results))],
        ["Categorical / Likert questions", str(n_cat)],
        ["Numeric questions", str(n_num)],
    ]
    summary_table = Table(summary_data, colWidths=[content_w * 0.55, content_w * 0.45])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, LIGHT_GREY]),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (-1, -1), TEXT_DARK),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("BOX", (0, 0), (-1, -1), 0.5, MID_GREY),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, MID_GREY),
    ]))
    story.append(summary_table)
    story.append(PageBreak())

    # ── Per-question pages ────────────────────────────────────────────────────
    for i, result in enumerate(results, start=1):
        q_label = result.label or result.variable
        heading = f"Q{i}. {q_label}"
        story.append(Paragraph(heading, styles["h1"]))
        story.append(Paragraph(
            f"<i>Variable:</i> <b>{result.variable}</b> &nbsp;|&nbsp; "
            f"<i>Type:</i> <b>{result.type.capitalize()}</b> &nbsp;|&nbsp; "
            f"<i>Valid n:</i> <b>{result.n_valid:,}</b>",
            styles["body"],
        ))
        story.append(Spacer(1, 3 * mm))

        # Chart
        chart_path = chart_for_question(result)
        if chart_path and chart_path.exists():
            img = Image(str(chart_path), width=content_w * 0.85, height=None)
            img.hAlign = "LEFT"
            story.append(img)
            story.append(Spacer(1, 2 * mm))

        # Frequency table (categorical) or stats table (numeric)
        if result.type == "categorical" and result.frequencies:
            table_data = [["Response", "N", "%"]]
            for row in result.frequencies:
                table_data.append([row.label, str(row.count), f"{row.percent:.1f}%"])

            freq_table = Table(
                table_data,
                colWidths=[content_w * 0.60, content_w * 0.18, content_w * 0.18],
            )
            freq_table.setStyle(_freq_table_style())
            story.append(freq_table)

        elif result.type == "numeric":
            stats_data = [
                ["Statistic", "Value"],
                ["Mean", f"{result.mean:.2f}"],
                ["Median", f"{result.median:.2f}"],
                ["Std Dev", f"{result.std:.2f}"],
                ["Min", f"{result.min:.2f}"],
                ["Max", f"{result.max:.2f}"],
            ]
            stats_table = Table(stats_data, colWidths=[content_w * 0.4, content_w * 0.3])
            stats_table.setStyle(_freq_table_style())
            story.append(stats_table)

        story.append(Spacer(1, 6 * mm))

        # Page break after every question except the last
        if i < len(results):
            story.append(PageBreak())

    doc.build(story)
    return output_path


def _freq_table_style() -> TableStyle:
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BRAND_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT_GREY]),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("TEXTCOLOR", (0, 1), (-1, -1), TEXT_DARK),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("BOX", (0, 0), (-1, -1), 0.5, MID_GREY),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, MID_GREY),
    ])
