"""Builds a 16:9 PowerPoint deck from the survey analysis (same charts as the PDF)."""
from __future__ import annotations

import io
from datetime import date

import matplotlib.pyplot as plt
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

from pdf_report import _short, chart_slides

ACCENT = RGBColor(0x2B, 0x6C, 0xB0)
INK = RGBColor(0x1A, 0x20, 0x2C)
MUTED = RGBColor(0x71, 0x80, 0x96)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)


def _text(slide, text, left, top, width, height, size=14, bold=False, color=INK, align=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.font.name = "Calibri"
        if align is not None:
            p.alignment = align
    return box


def _content_slide(prs, title, subtitle, footer):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()
    _text(slide, title, Inches(0.5), Inches(0.3), Inches(12), Inches(0.6), size=28, bold=True)
    if subtitle:
        _text(slide, subtitle, Inches(0.5), Inches(0.9), Inches(12), Inches(0.4), size=14, color=MUTED)
    _text(slide, footer, Inches(0.5), Inches(7.05), Inches(12), Inches(0.3), size=9, color=MUTED)
    return slide


def _add_figure(slide, fig, top=Inches(1.4), bottom_margin=Inches(0.55)):
    """Insert a matplotlib figure, scaled to fit the area below the title."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    w_in, h_in = fig.get_size_inches()
    plt.close(fig)
    buf.seek(0)
    max_w = SLIDE_W - Inches(1)
    max_h = SLIDE_H - top - bottom_margin
    scale = min(max_w / w_in, max_h / h_in)
    w, h = Emu(int(w_in * scale)), Emu(int(h_in * scale))
    slide.shapes.add_picture(buf, int((SLIDE_W - w) / 2), top, w, h)


def build_pptx(
    df: pd.DataFrame,
    report_title: str,
    summary: pd.DataFrame,
    cat_cols: list,
    filters_text: str,
    group: str | None = None,
    chosen: list | None = None,
) -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    footer = f"{report_title} · {date.today():%d %b %Y} · Aggregate results only, no individual data"

    # ---- Title slide ----
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = ACCENT
    _text(slide, report_title, Inches(0.8), Inches(2.4), Inches(11.5), Inches(1.2), size=40, bold=True, color=WHITE)
    _text(slide, f"{len(df)} responses", Inches(0.8), Inches(3.6), Inches(11.5), Inches(0.5), size=20, color=WHITE)
    _text(slide, filters_text, Inches(0.8), Inches(4.2), Inches(11.5), Inches(1), size=14, color=WHITE)
    _text(slide, f"Generated {date.today():%d %B %Y}", Inches(0.8), Inches(6.6), Inches(8), Inches(0.4),
          size=12, color=WHITE)

    # ---- At a glance ----
    top = summary.sort_values("% Yes", ascending=False)
    slide = _content_slide(prs, "At a glance", "Questions with the highest and lowest share of 'Yes' answers", footer)
    for col, (label, rows) in enumerate([("Highest % Yes", top.head(7)), ("Lowest % Yes", top.tail(7).iloc[::-1])]):
        x = Inches(0.5 + col * 6.3)
        _text(slide, label, x, Inches(1.45), Inches(6), Inches(0.4), size=16, bold=True, color=ACCENT)
        for i, (_, r) in enumerate(rows.iterrows()):
            y = Inches(1.95 + i * 0.7)
            _text(slide, f"{r['% Yes']:.0f}%", x, y, Inches(1.1), Inches(0.6), size=22, bold=True, color=ACCENT)
            _text(slide, _short(r["Question"], 60), x + Inches(1.15), y + Inches(0.02), Inches(4.9), Inches(0.65),
                  size=12)

    # ---- Chart slides (same as the PDF) ----
    for title, subtitle, fig in chart_slides(df, summary, cat_cols, group, chosen):
        slide = _content_slide(prs, title, subtitle, footer)
        _add_figure(slide, fig)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
