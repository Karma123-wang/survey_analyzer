"""Power BI-style charts and PDF report for the survey analysis.

chart_slides() draws every chart once; the PDF (here) and the PowerPoint (ppt_report.py) both use it.
"""
from __future__ import annotations

import io
import textwrap
from datetime import date

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

# ---------------- Power BI theme ----------------
BLUE = "#118DFF"
HIGHLIGHT = "#D64550"          # red bars for highlighted questions
NO_GREY = "#E6E6E6"
DARK = "#252423"               # title page
HEADER = "#118DFF"             # sky-blue Power BI header band on each page
HEADER_STRIP = "#12239E"       # navy accent strip at the left of the header
HEADER_SUBTEXT = "#E3F1FF"     # subtitle text on the header
CANVAS = "#F3F2F1"             # page background
INK = "#252423"
MUTED = "#605E5C"
GRID = "#EDEBE9"
PALETTE = ["#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7", "#744EC2", "#D9B300", "#D64550"]
ACCENT = BLUE                  # kept for compatibility

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "text.color": INK,
    "axes.edgecolor": GRID, "axes.labelcolor": MUTED, "xtick.color": MUTED, "ytick.color": INK,
    "axes.spines.top": False, "axes.spines.right": False, "figure.facecolor": "white",
})

PAGE_W, PAGE_H = landscape(A4)
MARGIN = 28


ALL_SECTIONS = ["summary", "notes", "focus", "profile", "yesno", "compare"]


def responses_line(df, filters_text: str) -> str:
    """'461 responses', plus the filter only when one is applied."""
    text = f"{len(df):,} responses"
    return f"{text}  ·  {filters_text}" if filters_text else text


def pct_yes(s: pd.Series) -> float:
    answered = s.isin(["Yes", "No"]).sum()
    return (s == "Yes").sum() / answered * 100 if answered else 0.0


def _short(text: str, width: int = 60) -> str:
    """Strip numbering like '2. ' and wrap long question text."""
    t = str(text).strip()
    if len(t) > 3 and t[0].isdigit() and "." in t[:4]:
        t = t.split(".", 1)[1].strip()
    t = " ".join(t.split())
    if len(t) > width * 2:
        t = t[: width * 2 - 1] + "…"
    return "\n".join(textwrap.wrap(t, width))


def _one_line(text: str) -> str:
    return _short(text, 1000).replace("\n", " ")


def _clean_axes(ax, xgrid=True):
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color(GRID)
    ax.tick_params(length=0)
    if xgrid:
        ax.xaxis.grid(True, color=GRID, linewidth=0.8)
        ax.set_axisbelow(True)



ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10, "XI": 11,
         "XII": 12, "PP": 0}


def _natural_key(v):
    """Sort classes (IX, X, XI…) and ages numerically, everything else alphabetically."""
    s = str(v).strip().upper()
    if s in ROMAN:
        return (0, ROMAN[s], s)
    try:
        return (0, float(s), s)
    except ValueError:
        return (1, 0, s)


def focus_groups(cat_cols):
    """Pick the Class, Age and Sex columns (in that order) for the red-indicator focus pages."""
    picked = []
    for words in (["class", "grade"], ["age"], ["sex", "gender"]):
        for c in cat_cols:
            if any(w in c.lower() for w in words) and c not in picked:
                picked.append(c)
                break
    return picked or list(cat_cols[:3])


def focus_figure(df, question_col, label, groups):
    """One panel per group column: % (and number) answering Yes in each group."""
    answered = df[question_col].isin(["Yes", "No"])
    n_yes = int((df[question_col] == "Yes").sum())
    fig, axes = plt.subplots(1, len(groups), figsize=(5.2 * len(groups), 5.6), squeeze=False)
    for ax, g in zip(axes[0], groups):
        sub = df[answered & df[g].notna()]
        tab = sub.groupby(g)[question_col].agg(n="size", yes=lambda s: (s == "Yes").sum())
        tab = tab[tab["n"] > 0]
        tab = tab.loc[sorted(tab.index, key=_natural_key)]
        pct = tab["yes"] / tab["n"] * 100
        names = [str(i) for i in tab.index][::-1]
        ax.barh(names, [100] * len(names), color=NO_GREY, height=0.6)
        ax.barh(names, pct.values[::-1], color=HIGHLIGHT, height=0.6)
        most = tab["yes"].max()
        for i, (y, n, p) in enumerate(zip(tab["yes"].values[::-1], tab["n"].values[::-1], pct.values[::-1])):
            small = n < 10
            ax.text(101, i, f"{y} of {n}  ({p:.0f}%)" + ("  small group" if small else ""), va="center",
                    fontsize=9.5 if not small else 8.5, color=MUTED if small else (HIGHLIGHT if y == most and y > 0 else INK),
                    fontweight="bold" if y == most and y > 0 else "normal")
        overall = n_yes / max(1, int(answered.sum())) * 100
        ax.axvline(overall, color=INK, linestyle="--", linewidth=1)
        ax.set_xlim(0, 150)
        ax.set_xticks([0, 25, 50, 75, 100])
        ax.xaxis.set_major_formatter(mticker.PercentFormatter())
        ax.set_title(f"By {g.rstrip(': ').lower()}", loc="left", fontsize=12, fontweight="bold", color=INK)
        _clean_axes(ax, xgrid=False)
        ax.tick_params(axis="y", labelsize=9.5)
    fig.text(0.01, 0.005, f"Red bar = % answering Yes in each group · dashed line = whole school ({overall:.0f}%) · "
                          f"bold red = most students · grey = fewer than 10 students (interpret with care)",
             fontsize=9, color=MUTED)
    fig.tight_layout(w_pad=2.5, rect=(0, 0.03, 1, 1))
    return fig, n_yes, int(answered.sum())

# ---------------- Charts (shared by PDF and PowerPoint) ----------------

def chart_slides(df, summary, cat_cols, group=None, chosen=None, sections=None):
    """Yield (title, subtitle, matplotlib figure) for each chart page/slide."""
    secs = set(sections or ALL_SECTIONS)
    label_of = dict(zip(summary["Original"], summary["Question"])) if "Original" in summary else {}
    # ---- Respondent profile: donut charts ----
    profile_cols = cat_cols[:6] if "profile" in secs else []
    if profile_cols:
        n = len(profile_cols)
        ncols = 3 if n > 2 else n
        nrows = (n + ncols - 1) // ncols
        fig, axes = plt.subplots(nrows, ncols, figsize=(5.2 * ncols, 3.9 * nrows), squeeze=False)
        for ax, col in zip(axes.flat, profile_cols):
            counts = df[col].value_counts().head(8)
            wedges, _ = ax.pie(counts.values, colors=PALETTE[: len(counts)], startangle=90, counterclock=False,
                               wedgeprops=dict(width=0.38, edgecolor="white", linewidth=1.5))
            ax.text(0, 0.06, f"{counts.sum()}", ha="center", va="center", fontsize=17, fontweight="bold")
            ax.text(0, -0.2, "responses", ha="center", va="center", fontsize=8, color=MUTED)
            ax.set_title(col.rstrip(": "), fontsize=12, fontweight="bold", loc="left", color=INK, pad=6)
            total = counts.sum()
            labels = [f"{_short(k, 20)}  {v / total * 100:.0f}%" for k, v in counts.items()]
            ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(0.98, 0.5), frameon=False,
                      fontsize=8.5, handlelength=0.9, handleheight=0.9)
            ax.set_aspect("equal")
        for ax in list(axes.flat)[n:]:
            ax.axis("off")
        fig.tight_layout(w_pad=3)
        yield "Respondent profile", f"{len(df)} responses", fig

    # ---- Focus pages: each red (key) indicator by class, age and sex ----
    if "focus" in secs and "Highlight" in summary:
        groups = focus_groups(cat_cols)
        for _, r in summary[summary["Highlight"]].iterrows():
            qcol = r["Original"] if "Original" in summary else r["Question"]
            if qcol not in df.columns or not groups:
                continue
            fig, n_yes, n_ans = focus_figure(df, qcol, r["Question"], groups)
            title = _one_line(r["Question"])
            fig.suptitle("\n".join(textwrap.wrap(f"“{title}”", 120)), x=0.01, ha="left", fontsize=12.5,
                         fontweight="bold", color=HIGHLIGHT)
            fig.tight_layout(w_pad=2.5, rect=(0, 0.03, 1, 0.9 if len(title) > 120 else 0.93))
            short = title if len(title) <= 58 else title[:57].rsplit(" ", 1)[0] + "…"
            yield ("Focus: " + short,
                   f"{n_yes} of {n_ans} students ({n_yes / max(1, n_ans) * 100:.0f}%) answered Yes · "
                   f"by {', '.join(g.rstrip(': ').lower() for g in groups)}", fig)

    # ---- Yes / No: 100% stacked bars ----
    if "yesno" not in secs or summary.empty:
        summary = summary.iloc[0:0]
    n_pages = -(-len(summary) // 12) if len(summary) else 0  # at most 12 questions per page, spread evenly
    per_page = max(1, -(-len(summary) // max(1, n_pages)))
    pages = [summary.iloc[i: i + per_page] for i in range(0, len(summary), per_page)]
    for p, chunk in enumerate(pages, 1):
        fig, ax = plt.subplots(figsize=(13, 0.56 * len(chunk) + 1.1))
        labels = [_short(q, 62) for q in chunk["Question"]][::-1]
        yes = chunk["% Yes"].values[::-1]
        red = list(chunk["Highlight"].values[::-1]) if "Highlight" in chunk else [False] * len(chunk)
        ax.barh(labels, yes, color=[HIGHLIGHT if r else BLUE for r in red], height=0.62)
        ax.barh(labels, 100 - yes, left=yes, color=NO_GREY, height=0.62)
        for i, v in enumerate(yes):
            if v >= 7:
                ax.text(1.2, i, f"{v:.0f}%", va="center", ha="left", fontsize=9.5, color="white", fontweight="bold")
            else:
                ax.text(v + 1, i, f"{v:.0f}%", va="center", ha="left", fontsize=9.5, color=INK, fontweight="bold")
            if 100 - v >= 7:
                ax.text(98.8, i, f"{100 - v:.0f}%", va="center", ha="right", fontsize=9, color=MUTED)
        ax.set_xlim(0, 100)
        ax.xaxis.set_major_formatter(mticker.PercentFormatter())
        ax.tick_params(axis="y", labelsize=9)
        _clean_axes(ax, xgrid=False)
        handles = [plt.Rectangle((0, 0), 1, 1, color=BLUE), plt.Rectangle((0, 0), 1, 1, color=NO_GREY)]
        legend_labels = ["Yes", "No"]
        if any(red):
            handles.insert(1, plt.Rectangle((0, 0), 1, 1, color=HIGHLIGHT))
            legend_labels.insert(1, "Yes (key indicator)")
        ax.legend(handles, legend_labels, loc="lower left", bbox_to_anchor=(0, 1.0), ncol=3, frameon=False,
                  fontsize=9, handlelength=0.9)
        fig.tight_layout()
        yield "Yes / No results", f"Part {p} of {len(pages)} · % of respondents answering Yes and No", fig

    # ---- Group comparison: clustered horizontal bars ----
    if group and chosen and "compare" in secs:
        table = df.groupby(group)[chosen].agg(pct_yes)
        sizes = df.groupby(group).size()
        groups = list(table.index)
        g = max(1, len(groups))
        fig, ax = plt.subplots(figsize=(13, max(3.6, len(chosen) * (0.32 * g + 0.35) + 1.2)))
        bar_h = 0.78 / g
        for gi, grp in enumerate(groups):
            ys = [i + gi * bar_h for i in range(len(chosen))]
            vals = [table.loc[grp, q] for q in chosen]
            ax.barh(ys, vals, height=bar_h * 0.86, color=PALETTE[gi % len(PALETTE)], label=f"{grp} (n={sizes[grp]})")
            for y, v in zip(ys, vals):
                ax.text(v + 0.8, y, f"{v:.0f}%", va="center", fontsize=8.5, color=INK)
        ax.set_yticks([i + bar_h * (g - 1) / 2 for i in range(len(chosen))])
        ax.set_yticklabels([_short(label_of.get(q, q), 55) for q in chosen], fontsize=9)
        ax.invert_yaxis()
        ax.set_xlim(0, 108)
        ax.xaxis.set_major_formatter(mticker.PercentFormatter())
        _clean_axes(ax)
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=min(g, 6), frameon=False, fontsize=9,
                  handlelength=0.9)
        fig.tight_layout()
        yield f"Comparison by {group.rstrip(': ')}", "% answering Yes in each group", fig


def kpi_items(df, summary):
    """KPI cards: responses, questions, and each highlighted (red) question."""
    items = [(f"{len(df):,}", "Responses", False), (str(len(summary)), "Yes / No questions", False)]
    if "Highlight" in summary:
        for _, r in summary[summary["Highlight"]].head(3).iterrows():
            items.append((f"{r['% Yes']:.0f}%", _one_line(r["Question"]), True))
    return items


# ---------------- PDF drawing helpers ----------------

def _canvas_bg(c):
    c.setFillColor(CANVAS)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)


def _header(c, title, subtitle):
    c.setFillColor(HEADER)
    c.rect(0, PAGE_H - 62, PAGE_W, 62, fill=1, stroke=0)
    c.setFillColor(HEADER_STRIP)
    c.rect(0, PAGE_H - 62, 6, 62, fill=1, stroke=0)
    c.setFillColor("white")
    c.setFont("Helvetica-Bold", 20)
    c.drawString(MARGIN, PAGE_H - 34, title)
    if subtitle:
        c.setFillColor(HEADER_SUBTEXT)
        c.setFont("Helvetica", 10.5)
        c.drawString(MARGIN, PAGE_H - 51, subtitle)


def _footer(c, footer):
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8.5)
    c.drawString(MARGIN, 14, footer)
    c.drawRightString(PAGE_W - MARGIN, 14, f"Page {c.getPageNumber()}")


def _card(c, x, y, w, h, top_colour=None):
    c.setFillColor("#E1DFDD")
    c.roundRect(x + 1, y - 1.5, w, h, 5, fill=1, stroke=0)  # soft shadow
    c.setFillColor("white")
    c.roundRect(x, y, w, h, 5, fill=1, stroke=0)
    if top_colour:
        c.setFillColor(top_colour)
        c.rect(x + 2, y + h - 4, w - 4, 4, fill=1, stroke=0)


def _fig_image(fig) -> ImageReader:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)


def _image_in_card(c, img):
    x, y, w, h = MARGIN, 32, PAGE_W - 2 * MARGIN, PAGE_H - 62 - 16 - 32
    _card(c, x, y, w, h)
    iw, ih = img.getSize()
    pad = 14
    scale = min((w - 2 * pad) / iw, (h - 2 * pad) / ih)
    dw, dh = iw * scale, ih * scale
    c.drawImage(img, x + (w - dw) / 2, y + h - pad - dh, dw, dh)


def _wrap_lines(text, width, max_lines):
    lines = textwrap.wrap(text, max(10, width))
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1][:-1] + "…"
    return lines


# ---------------- PDF ----------------

def build_pdf(
    df: pd.DataFrame,
    report_title: str,
    summary: pd.DataFrame,
    cat_cols: list,
    filters_text: str,
    group: str | None = None,
    chosen: list | None = None,
    notes: dict | None = None,
    sections: list | None = None,
) -> bytes:
    """notes = {"findings": [...], "recommendations": [...]}; sections = pages to include."""
    secs = set(sections or ALL_SECTIONS)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=landscape(A4))
    parts = [x.strip() for x in report_title.split("\n")] + ["", ""]
    main_title = parts[0] or "Survey Report"
    sub_title, author = parts[1], parts[2]                   # e.g. "DCS - Counsellor", "Seldon Lhamo"
    footer = " – ".join(x for x in (main_title, sub_title) if x) + " · Aggregate results only, no individual data"

    # ---- Premium cover page: title at the top, author block at the bottom ----
    c.setFillColor(HEADER)                                   # sky-blue background
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.saveState()                                            # soft decorative circles on the right
    c.setFillColor("white")
    for (cx, cy, r, a) in [(PAGE_W - 30, PAGE_H - 20, 240, 0.07), (PAGE_W - 120, 40, 170, 0.06),
                           (PAGE_W - 250, PAGE_H / 2 - 10, 90, 0.05)]:
        c.setFillAlpha(a)
        c.circle(cx, cy, r, fill=1, stroke=0)
    c.restoreState()
    c.setFillColor(HEADER_STRIP)                             # navy strip
    c.rect(0, 0, 16, PAGE_H, fill=1, stroke=0)

    x0 = 64
    # Top: label, accent bar, title
    c.setFillColor("#CFE6FF")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(x0, PAGE_H - 70, "R E P O R T")
    c.setFillColor("white")
    c.rect(x0, PAGE_H - 92, 64, 4, fill=1, stroke=0)
    title_size = 40 if len(main_title) <= 36 else 34
    y = PAGE_H - 92 - 22 - title_size
    c.setFont("Helvetica-Bold", title_size)
    for row in _wrap_lines(main_title, 32 if title_size == 40 else 38, 3):
        c.drawString(x0, y, row)
        y -= title_size + 8
    c.setFillColor("#E3F1FF")
    c.setFont("Helvetica", 13)
    c.drawString(x0, y - 4, responses_line(df, filters_text)[:120])

    # Bottom: divider, PREPARED BY, name, role
    base = 70
    c.setStrokeColor("white")
    c.setStrokeAlpha(0.55)
    c.setLineWidth(0.8)
    c.line(x0, base + 92, x0 + 300, base + 92)
    c.setStrokeAlpha(1)
    if author or sub_title:
        c.setFillColor("#CFE6FF")
        c.setFont("Helvetica-Bold", 9.5)
        c.drawString(x0, base + 70, "P R E P A R E D   B Y")
    if author:
        c.setFillColor("white")
        c.setFont("Helvetica-Bold", 22)
        c.drawString(x0, base + 40, author)
    if sub_title:
        c.setFillColor("#E3F1FF")
        c.setFont("Helvetica", 14)
        c.drawString(x0, base + (16 if author else 40), sub_title)
    c.showPage()

    if "summary" in secs:
        # ---- Summary page: KPI cards + highest / lowest ----
        _canvas_bg(c)
        _header(c, "Summary", "Key indicators and the questions with the highest and lowest % Yes")
        items = kpi_items(df, summary)
        gap = 12
        kw = (PAGE_W - 2 * MARGIN - gap * (len(items) - 1)) / len(items)
        kh = 92
        ky = PAGE_H - 62 - 16 - kh
        for i, (value, label, red) in enumerate(items):
            x = MARGIN + i * (kw + gap)
            _card(c, x, ky, kw, kh, top_colour=HIGHLIGHT if red else BLUE)
            c.setFillColor(HIGHLIGHT if red else INK)
            c.setFont("Helvetica-Bold", 28)
            c.drawString(x + 14, ky + kh - 42, value)
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 9)
            for j, line in enumerate(_wrap_lines(label, int(kw / 4.6), 3)):
                c.drawString(x + 14, ky + kh - 58 - j * 11, line)

        top = summary.sort_values("% Yes", ascending=False)
        ch_y = 32
        ch_h = ky - 16 - ch_y
        cw = (PAGE_W - 2 * MARGIN - gap) / 2
        for col, (label, rows) in enumerate([("Highest % Yes", top.head(7)), ("Lowest % Yes", top.tail(7).iloc[::-1])]):
            x = MARGIN + col * (cw + gap)
            _card(c, x, ch_y, cw, ch_h)
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x + 14, ch_y + ch_h - 24, label)
            row_h = (ch_h - 44) / 7
            for i, (_, r) in enumerate(rows.iterrows()):
                yy = ch_y + ch_h - 50 - i * row_h
                red = bool(r.get("Highlight", False))
                bar_w = 70
                c.setFillColor(NO_GREY)
                c.rect(x + 14, yy - 9, bar_w, 6, fill=1, stroke=0)
                c.setFillColor(HIGHLIGHT if red else BLUE)
                c.rect(x + 14, yy - 9, bar_w * r["% Yes"] / 100, 6, fill=1, stroke=0)
                c.setFillColor(HIGHLIGHT if red else INK)
                c.setFont("Helvetica-Bold", 12)
                c.drawString(x + 14, yy + 1, f"{r['% Yes']:.0f}%")
                c.setFillColor(INK)
                c.setFont("Helvetica", 9)
                for j, line in enumerate(_wrap_lines(_one_line(r["Question"]), int((cw - 120) / 4.4), 2)):
                    c.drawString(x + 98, yy + 1 - j * 10.5, line)
        _footer(c, footer)
        c.showPage()


    # ---- Key findings & recommendations ----
    if "notes" in secs and notes and (notes.get("findings") or notes.get("recommendations")):
        _canvas_bg(c)
        _header(c, "Key findings & recommendations", sub_title or "")
        boxes = [(t, items) for t, items in (("Key findings", notes.get("findings") or []),
                                               ("Recommendations", notes.get("recommendations") or [])) if items]
        gap = 12
        bw = (PAGE_W - 2 * MARGIN - gap * (len(boxes) - 1)) / len(boxes)
        by, bh = 32, PAGE_H - 62 - 16 - 32
        for bi, (heading, items) in enumerate(boxes):
            x = MARGIN + bi * (bw + gap)
            _card(c, x, by, bw, bh, top_colour=BLUE if bi == 0 else HEADER_STRIP)
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", 15)
            c.drawString(x + 18, by + bh - 34, heading)
            yy = by + bh - 62
            chars = int((bw - 60) / 6.7)
            for n, item in enumerate(items, 1):
                lines = textwrap.wrap(item, chars) or [""]
                if yy - 19 * len(lines) < by + 14:
                    break
                c.setFillColor(BLUE if bi == 0 else HEADER_STRIP)
                c.setFont("Helvetica-Bold", 13.5)
                c.drawString(x + 18, yy, f"{n}." if bi == 1 else "•")
                c.setFillColor(INK)
                c.setFont("Helvetica", 13.5)
                for line in lines:
                    c.drawString(x + 38, yy, line)
                    yy -= 19
                yy -= 10
        _footer(c, footer)
        c.showPage()

    # ---- Chart pages ----
    for title, subtitle, fig in chart_slides(df, summary, cat_cols, group, chosen, secs):
        _canvas_bg(c)
        _header(c, title, subtitle)
        _image_in_card(c, _fig_image(fig))
        _footer(c, footer)
        c.showPage()

    c.save()
    return buf.getvalue()
