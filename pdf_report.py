"""Builds a landscape, presentation-style PDF report from the survey analysis."""
from __future__ import annotations

import io
import textwrap
from datetime import date

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

PAGE_W, PAGE_H = landscape(A4)
MARGIN = 40
ACCENT = "#2b6cb0"
INK = "#1a202c"
MUTED = "#718096"
PALETTE = ["#2b6cb0", "#dd6b20", "#38a169", "#805ad5", "#d53f8c", "#319795", "#b7791f", "#e53e3e"]


def _short(text: str, width: int = 60) -> str:
    """Strip numbering like '2. ' and wrap long question text."""
    t = str(text).strip()
    if len(t) > 3 and t[0].isdigit() and "." in t[:4]:
        t = t.split(".", 1)[1].strip()
    t = " ".join(t.split())
    if len(t) > width * 2:
        t = t[: width * 2 - 1] + "…"
    return "\n".join(textwrap.wrap(t, width))


def _fig_to_image(fig) -> ImageReader:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)


def _page(c: canvas.Canvas, title: str, subtitle: str, footer: str):
    c.setFillColor(ACCENT)
    c.rect(0, PAGE_H - 8, PAGE_W, 8, fill=1, stroke=0)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(MARGIN, PAGE_H - 50, title)
    if subtitle:
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 11)
        c.drawString(MARGIN, PAGE_H - 70, subtitle)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9)
    c.drawString(MARGIN, 20, footer)
    c.drawRightString(PAGE_W - MARGIN, 20, f"Page {c.getPageNumber()}")


def _draw_image(c: canvas.Canvas, img: ImageReader, top_offset: float = 85):
    """Fit the image into the content area below the title."""
    iw, ih = img.getSize()
    max_w = PAGE_W - 2 * MARGIN
    max_h = PAGE_H - top_offset - 40
    scale = min(max_w / iw, max_h / ih)
    w, h = iw * scale, ih * scale
    c.drawImage(img, (PAGE_W - w) / 2, PAGE_H - top_offset - h, w, h)


def pct_yes(s: pd.Series) -> float:
    answered = s.isin(["Yes", "No"]).sum()
    return (s == "Yes").sum() / answered * 100 if answered else 0.0


def chart_slides(df, summary, cat_cols, group=None, chosen=None):
    """Yield (title, subtitle, matplotlib figure) for each chart slide/page.
    Shared by the PDF and PowerPoint reports so both always match."""
    profile_cols = cat_cols[:6]
    if profile_cols:
        n = len(profile_cols)
        ncols = 3 if n > 2 else n
        nrows = (n + ncols - 1) // ncols
        fig, axes = plt.subplots(nrows, ncols, figsize=(5 * ncols, 3.4 * nrows), squeeze=False)
        for ax, col in zip(axes.flat, profile_cols):
            counts = df[col].value_counts().head(8)
            ax.barh([_short(i, 22) for i in counts.index[::-1]], counts.values[::-1], color=ACCENT)
            ax.set_title(col.rstrip(": "), fontsize=12, loc="left", color=INK)
            for s in ["top", "right"]:
                ax.spines[s].set_visible(False)
            for i, v in enumerate(counts.values[::-1]):
                ax.text(v, i, f" {v}", va="center", fontsize=9, color=MUTED)
            ax.tick_params(labelsize=9)
        for ax in list(axes.flat)[n:]:
            ax.axis("off")
        fig.tight_layout()
        yield "Respondent profile", f"{len(df)} responses", fig

    n_pages = max(1, -(-len(summary) // 12))  # at most 12 questions per page, spread evenly
    per_page = max(1, -(-len(summary) // n_pages))
    pages = [summary.iloc[i : i + per_page] for i in range(0, len(summary), per_page)]
    for p, chunk in enumerate(pages, 1):
        fig, ax = plt.subplots(figsize=(13, 0.55 * len(chunk) + 1))
        labels = [_short(q, 70) for q in chunk["Question"]][::-1]
        vals = chunk["% Yes"].values[::-1]
        ax.barh(labels, vals, color=ACCENT)
        ax.barh(labels, 100 - vals, left=vals, color="#e2e8f0")
        for i, v in enumerate(vals):
            ax.text(v + 1 if v < 85 else v - 1, i, f"{v:.0f}%", va="center",
                    ha="left" if v < 85 else "right", fontsize=10,
                    color=INK if v < 85 else "white", fontweight="bold")
        ax.set_xlim(0, 100)
        ax.set_xlabel("% answering Yes")
        ax.tick_params(axis="y", labelsize=9)
        for s in ["top", "right"]:
            ax.spines[s].set_visible(False)
        fig.tight_layout()
        yield "Yes / No questions", f"Part {p} of {len(pages)} · % of respondents answering Yes", fig

    if group and chosen:
        table = df.groupby(group)[chosen].agg(pct_yes)
        sizes = df.groupby(group).size()
        groups = list(table.index)
        fig, ax = plt.subplots(figsize=(13, max(3.5, 0.5 * len(chosen) * max(1, len(groups) / 2) + 1)))
        bar_h = 0.8 / max(1, len(groups))
        ypos = range(len(chosen))
        for gi, g in enumerate(groups):
            vals = [table.loc[g, q] for q in chosen]
            ax.barh([y + gi * bar_h for y in ypos], vals, height=bar_h,
                    color=PALETTE[gi % len(PALETTE)], label=f"{g} (n={sizes[g]})")
        ax.set_yticks([y + bar_h * (len(groups) - 1) / 2 for y in ypos])
        ax.set_yticklabels([_short(q, 60) for q in chosen], fontsize=9)
        ax.invert_yaxis()
        ax.set_xlim(0, 100)
        ax.set_xlabel("% answering Yes")
        ax.legend(loc="lower right", fontsize=9, frameon=False)
        for s in ["top", "right"]:
            ax.spines[s].set_visible(False)
        fig.tight_layout()
        yield f"Comparison by {group.rstrip(': ')}", "% answering Yes in each group", fig


def build_pdf(
    df: pd.DataFrame,
    report_title: str,
    summary: pd.DataFrame,
    cat_cols: list,
    filters_text: str,
    group: str | None = None,
    chosen: list | None = None,
) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=landscape(A4))
    footer = f"{report_title} · Generated {date.today():%d %b %Y} · Aggregate results only, no individual data"

    # ---- Title page ----
    c.setFillColor(ACCENT)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor("white")
    c.setFont("Helvetica-Bold", 34)
    c.drawString(60, PAGE_H / 2 + 40, report_title)
    c.setFont("Helvetica", 16)
    c.drawString(60, PAGE_H / 2, f"{len(df)} responses")
    c.setFont("Helvetica", 12)
    for i, line in enumerate(textwrap.wrap(filters_text, 110)[:4]):
        c.drawString(60, PAGE_H / 2 - 30 - i * 16, line)
    c.drawString(60, 60, f"Generated {date.today():%d %B %Y}")
    c.showPage()

    # ---- Key findings: top 'Yes' rates ----
    top = summary.sort_values("% Yes", ascending=False)
    _page(c, "At a glance", "Questions with the highest and lowest share of 'Yes' answers", footer)
    y = PAGE_H - 110
    col_x = [MARGIN, PAGE_W / 2 + 10]
    for col, (label, rows) in enumerate([("Highest % Yes", top.head(8)), ("Lowest % Yes", top.tail(8).iloc[::-1])]):
        x = col_x[col]
        c.setFillColor(ACCENT)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x, y, label)
        yy = y - 26
        for _, r in rows.iterrows():
            c.setFillColor(ACCENT)
            c.setFont("Helvetica-Bold", 16)
            c.drawRightString(x + 55, yy, f"{r['% Yes']:.0f}%")
            c.setFillColor(INK)
            c.setFont("Helvetica", 10)
            for j, line in enumerate(_short(r["Question"], 62).split("\n")[:2]):
                c.drawString(x + 65, yy + 2 - j * 12, line)
            yy -= 44
    c.showPage()

    # ---- Respondent profile ----
    for title, subtitle, fig in chart_slides(df, summary, cat_cols, group, chosen):
        _page(c, title, subtitle, footer)
        _draw_image(c, _fig_to_image(fig))
        c.showPage()

    c.save()
    return buf.getvalue()
