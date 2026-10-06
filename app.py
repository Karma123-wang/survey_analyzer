"""
Survey Analyser - Power BI-style dashboard for survey Excel / CSV exports.
Run:  pip install streamlit pandas openpyxl plotly matplotlib reportlab python-pptx
      streamlit run app.py
"""
import json

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from pdf_report import build_pdf
from ppt_report import build_pptx
from html_report import build_html

st.set_page_config(page_title="Survey Analyser", page_icon="📊", layout="wide")

# ---------------- Colours (Power BI default theme) ----------------
BLUE, RED, NO_GREY = "#118DFF", "#D64550", "#E6E6E6"
PBI_COLOURS = ["#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7", "#744EC2", "#D9B300", "#D64550"]
FONT = "Segoe UI, Helvetica Neue, Arial, sans-serif"

# Bars for questions containing these words are shown in red.
RED_DEFAULT_KEYWORDS = ["harsh form of punishment", "danger to myself"]

# ---------------- Page styling ----------------
st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{ background: #F3F2F1; }}
  .block-container {{ padding-top: 1.2rem; }}
  div[class*="st-key-card"] {{
      background: #FFFFFF; border-radius: 6px; padding: 14px 16px 6px 16px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.12); }}
  .pbi-header {{ background: {BLUE}; color: #FFFFFF; padding: 14px 22px; border-radius: 6px;
      border-left: 6px solid #12239E;
      margin-bottom: 14px; font-family: {FONT}; }}
  .pbi-header h1 {{ font-size: 24px; margin: 0; color: #FFFFFF; font-weight: 600; }}
  .pbi-header p {{ margin: 2px 0 0 0; color: #E3F1FF; font-size: 13px; }}
  .kpi {{ background: #FFFFFF; border-radius: 6px; padding: 14px 16px; height: 118px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.12); border-top: 4px solid {BLUE}; font-family: {FONT}; }}
  .kpi.red {{ border-top-color: {RED}; }}
  .kpi .value {{ font-size: 34px; font-weight: 600; color: #252423; line-height: 1.1; }}
  .kpi.red .value {{ color: {RED}; }}
  .kpi .label {{ font-size: 12.5px; color: #605E5C; margin-top: 6px; line-height: 1.25; }}
  .card-title {{ font-family: {FONT}; font-size: 15px; font-weight: 600; color: #252423; margin-bottom: 0; }}
</style>
""", unsafe_allow_html=True)


# ---------------- Helpers ----------------
PRIVATE_EXACT = {"name", "name:", "timestamp", "si.no", "sl.no", "cid", "phone", "email"}
PRIVATE_WORDS = ["address", "please write the name"]


def is_private(col: str) -> bool:
    c = col.lower().strip()
    return c in PRIVATE_EXACT or any(w in c for w in PRIVATE_WORDS)


def is_yes_no(series: pd.Series) -> bool:
    vals = set(series.dropna().astype(str).str.strip().unique())
    return 0 < len(vals) <= 2 and vals <= {"Yes", "No"}


def pct_yes(s: pd.Series) -> float:
    """% Yes among people who answered (blank answers are ignored)."""
    answered = s.isin(["Yes", "No"]).sum()
    return (s == "Yes").sum() / answered * 100 if answered else 0.0


# Written-answer questions (never treated as groups, even in small surveys)
TEXT_WORDS = ["reason", "specify", "please write", "write the name", "if yes", "if you selected"]


def is_category(series: pd.Series) -> bool:
    if any(w in str(series.name).lower() for w in TEXT_WORDS):
        return False
    return 1 < series.nunique() <= 12 and not is_yes_no(series)


def short(q: str, n: int = 70) -> str:
    """Remove question numbers like '2. ' and shorten long text for chart labels."""
    t = " ".join(str(q).split())
    if len(t) > 3 and t[0].isdigit() and "." in t[:4]:
        t = t.split(".", 1)[1].strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def pbi_style(fig, height=320):
    fig.update_layout(
        height=height, font=dict(family=FONT, size=12, color="#252423"),
        paper_bgcolor="white", plot_bgcolor="white", margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="h", yanchor="top", y=-0.08, x=0, title_text=""),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#EDEBE9", zeroline=False, title_text="")
    fig.update_yaxes(showgrid=False, title_text="")
    return fig


def card_title(text):
    st.markdown(f'<p class="card-title">{text}</p>', unsafe_allow_html=True)


def kpi(col, value, label, red=False):
    col.markdown(f'<div class="kpi {"red" if red else ""}"><div class="value">{value}</div>'
                 f'<div class="label">{label}</div></div>', unsafe_allow_html=True)


# ---------------- Data source (sidebar) ----------------
st.sidebar.header("Data")
file = st.sidebar.file_uploader("Upload survey file (.xlsx or .csv)", type=["xlsx", "csv"])
if not file:
    st.markdown('<div class="pbi-header"><h1>Survey Analyser</h1>'
                '<p>Upload a survey file in the sidebar to build the dashboard</p></div>', unsafe_allow_html=True)
    st.info("⬅️ Upload an Excel or CSV survey export in the sidebar to begin.")
    st.stop()

if file.name.endswith(".csv"):
    df = pd.read_csv(file)
    report_name = file.name.rsplit(".", 1)[0]
else:
    xls = pd.ExcelFile(file)
    sheet = st.sidebar.selectbox("Sheet", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet)
    report_name = sheet

df = df.dropna(how="all")
df.columns = [str(c).strip() for c in df.columns]
for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.strip().replace({"nan": None})

usable = [c for c in df.columns if not is_private(c)]
yn_cols = [c for c in usable if is_yes_no(df[c])]
cat_cols = [c for c in usable if is_category(df[c])]
total_rows = len(df)

# ---------------- Filters (slicers) ----------------
st.sidebar.header("Filters")
active_filters = []
for c in cat_cols[:6]:
    opts = sorted(df[c].dropna().unique())
    pick = st.sidebar.multiselect(c.rstrip(": "), opts, default=opts)
    df = df[df[c].isin(pick)]
    if len(pick) < len(opts):
        active_filters.append(f"{c.rstrip(': ')}: {', '.join(map(str, pick))}")
filters_text = "Filters: " + "; ".join(active_filters) if active_filters else ""

st.markdown(f'<div class="pbi-header"><h1>Situation Analysis &amp; Survey Report</h1>'
            f'<div style="font-size:18px;font-weight:600;margin-top:2px">{report_name} - Counsellor</div>'
            f'<p>{len(df):,} responses{" · " + filters_text if filters_text else ""}</p></div>',
            unsafe_allow_html=True)

if df.empty:
    st.warning("No responses match the filters.")
    st.stop()

# ---------------- Summary table ----------------
summary = pd.DataFrame({
    "Question": yn_cols,
    "% Yes": [pct_yes(df[c]) for c in yn_cols],
    "Yes": [(df[c] == "Yes").sum() for c in yn_cols],
    "No": [(df[c] == "No").sum() for c in yn_cols],
}).round(1)
summary["Highlight"] = summary["Question"].str.lower().apply(lambda q: any(k in q for k in RED_DEFAULT_KEYWORDS))

# ---------------- KPI cards ----------------
red_rows = summary[summary["Highlight"]]
kpi_cols = st.columns(2 + max(len(red_rows), 1))
kpi(kpi_cols[0], f"{len(df):,}", f"Responses<br>(of {total_rows:,} in file)")
kpi(kpi_cols[1], len(yn_cols), "Yes / No questions<br>analysed")
if len(red_rows):
    for col, (_, r) in zip(kpi_cols[2:], red_rows.iterrows()):
        kpi(col, f"{r['% Yes']:.0f}%", short(r["Question"], 75), red=True)
else:
    kpi(kpi_cols[2], f"{summary['% Yes'].mean():.0f}%" if len(summary) else "–", "Average % Yes")
st.write("")

# ---------------- Pages (like Power BI report pages) ----------------
tab_overview, tab_yesno, tab_compare, tab_text, tab_export = st.tabs(
    ["Overview", "Yes / No results", "Compare groups", "Written reasons", "✏️ Edit & export"])

# ---- Overview: donut charts ----
with tab_overview:
    for row_start in range(0, len(cat_cols), 3):
        cols = st.columns(3)
        for i, c in enumerate(cat_cols[row_start:row_start + 3]):
            with cols[i]:
                with st.container(key=f"card_donut_{row_start + i}"):
                    card_title(c.rstrip(": "))
                    counts = df[c].value_counts().reset_index()
                    counts.columns = ["Value", "Count"]
                    fig = px.pie(counts, names="Value", values="Count", hole=0.6,
                                 color_discrete_sequence=PBI_COLOURS)
                    fig.update_traces(textinfo="percent", textposition="outside", sort=False,
                                      hovertemplate="%{label}: %{value} (%{percent})<extra></extra>")
                    fig.add_annotation(text=f"<b>{counts['Count'].sum()}</b>", showarrow=False,
                                       font=dict(size=20, family=FONT))
                    st.plotly_chart(pbi_style(fig, 300), use_container_width=True, key=f"donut_{row_start + i}")
        st.write("")

# ---- Yes / No results: stacked Yes vs No bars ----
with tab_yesno:
    with st.container(key="card_yesno"):
        card_title("% of respondents answering Yes / No")
        labels = [short(q) for q in summary["Question"]]
        yes = summary["% Yes"].values
        no = 100 - yes
        fig = go.Figure()
        fig.add_bar(y=labels, x=yes, orientation="h", name="Yes",
                    marker_color=[RED if h else BLUE for h in summary["Highlight"]],
                    text=[f"{v:.0f}%" for v in yes], textposition="inside", insidetextanchor="start",
                    textfont=dict(color="white"), hovertemplate="%{y}<br>Yes: %{x:.0f}%<extra></extra>")
        fig.add_bar(y=labels, x=no, orientation="h", name="No", marker_color=NO_GREY,
                    text=[f"{v:.0f}%" for v in no], textposition="inside", insidetextanchor="end",
                    textfont=dict(color="#605E5C"), hovertemplate="%{y}<br>No: %{x:.0f}%<extra></extra>")
        fig.update_layout(barmode="stack", bargap=0.25, showlegend=True)
        fig.update_yaxes(autorange="reversed", categoryorder="array", categoryarray=labels)
        fig.update_xaxes(range=[0, 100], ticksuffix="%")
        st.plotly_chart(pbi_style(fig, 34 * len(labels) + 80), use_container_width=True, key="chart_yesno")
    st.write("")
    with st.container(key="card_table"):
        card_title("Table")
        st.dataframe(summary.drop(columns="Highlight"), use_container_width=True, hide_index=True,
                     column_config={"% Yes": st.column_config.ProgressColumn(
                         "% Yes", format="%.0f%%", min_value=0, max_value=100)})

# ---- Compare groups: clustered bars ----
group, chosen = None, []
with tab_compare:
    if cat_cols and yn_cols:
        c1, c2 = st.columns([1, 3])
        group = c1.selectbox("Break down by", cat_cols, format_func=lambda c: c.rstrip(": "))
        chosen = c2.multiselect("Questions", yn_cols, default=yn_cols[:5], format_func=short)
        if chosen:
            table = df.groupby(group)[chosen].agg(pct_yes).round(0)
            sizes = df.groupby(group).size()
            long = table.reset_index().melt(id_vars=group, var_name="Question", value_name="% Yes")
            long["Question"] = long["Question"].map(lambda q: short(q, 60))
            long[group] = long[group].astype(str) + " (n=" + long[group].map(sizes).astype(str) + ")"
            with st.container(key="card_compare"):
                card_title(f"% answering Yes by {group.rstrip(': ')}")
                n_groups = long[group].nunique()
                fig = px.bar(long, y="Question", x="% Yes", color=group, barmode="group", orientation="h",
                             text="% Yes", color_discrete_sequence=PBI_COLOURS,
                             category_orders={"Question": list(dict.fromkeys(long["Question"]))})
                fig.update_traces(texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False)
                fig.update_layout(bargap=0.3, bargroupgap=0.12)  # space between questions and between bars
                fig.update_xaxes(range=[0, 112], ticksuffix="%")
                fig.update_yaxes(autorange="reversed")  # first question at the top
                fig = pbi_style(fig, max(320, len(chosen) * (28 * n_groups + 40) + 90))
                fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title_text=""),
                                  margin=dict(l=10, r=30, t=40, b=10))
                st.plotly_chart(fig, use_container_width=True, key="chart_compare")

# ---- Written reasons ----
with tab_text:
    text_cols = [c for c in usable if c not in yn_cols and c not in cat_cols
                 and (df[c].nunique() > 12 or any(w in c.lower() for w in TEXT_WORDS))]
    if text_cols:
        with st.container(key="card_text"):
            tc = st.selectbox("Question", text_cols, format_func=lambda c: short(c, 100))
            answers = df[tc].dropna()
            answers = answers[answers.str.len() > 3]  # skip "-", "no", etc.
            keyword = st.text_input("Count answers containing a keyword (e.g. study, hostel, exam)")
            if keyword:
                hits = answers[answers.str.lower().str.contains(keyword.lower(), regex=False)]
                st.markdown(f"**{len(hits)}** of {len(answers)} answers mention “{keyword}”")
                answers = hits
            st.dataframe(answers.reset_index(drop=True), use_container_width=True)
    else:
        st.info("No written-answer columns found.")

# ---- Edit & export ----
SECTION_LABELS = {"summary": "Summary", "notes": "Key findings & recommendations", "profile": "Respondent profile",
                  "yesno": "Yes / No results", "compare": "Comparison"}


def clean_question(q: str) -> str:
    t = " ".join(str(q).split())
    if len(t) > 3 and t[0].isdigit() and "." in t[:4]:
        t = t.split(".", 1)[1].strip()
    return t


def lines_of(text: str) -> list:
    return [ln.strip(" •-*\t") for ln in str(text).splitlines() if ln.strip(" •-*\t")]


def auto_findings(df, summary, group, chosen) -> str:
    """Draft key findings from the data (the user can edit them)."""
    out = [f"{len(df):,} students responded to the survey."]
    for _, r in summary[summary["Highlight"]].iterrows():
        out.append(f"{r['% Yes']:.0f}% answered Yes to: \"{clean_question(r['Question'])}\"")
    if group and chosen:
        table = df.groupby(group)[chosen].agg(pct_yes)
        gaps = []
        for q in chosen:
            col = table[q].dropna()
            if len(col) >= 2:
                gaps.append((col.max() - col.min(), q, col.idxmax(), col.max(), col.idxmin(), col.min()))
        for gap, q, hi, hv, lo, lv in sorted(gaps, reverse=True)[:3]:
            if gap >= 10:
                out.append(f"\"{clean_question(q)}\": {hi} {hv:.0f}% vs {lo} {lv:.0f}% "
                           f"({group.rstrip(': ')} difference of {gap:.0f} points).")
    return "\n".join(out)


ss = st.session_state
# Defaults (the role line follows the chosen sheet until you change it)
ss.setdefault("ed_title", "Situation Analysis & Survey Report")
ss.setdefault("ed_author", "Seldon Lhamo")
ss.setdefault("ed_findings", "")
ss.setdefault("ed_recs", "")
ss.setdefault("ed_sections", list(SECTION_LABELS))
ss.setdefault("q_settings", {})          # original question -> {"show", "label", "red"}
if ss.get("ed_sheet") != report_name:
    ss["ed_role"] = f"{report_name} - Counsellor"
    ss["ed_sheet"] = report_name

with tab_export:
    suggested = auto_findings(df, summary, group, chosen)
    if ss.pop("refill_findings", False):            # "Refill" button was pressed
        ss["findings_is_auto"] = True
        ss["ed_findings"] = ""
    # Keep the automatic draft up to date only until the user writes their own text
    if ss.get("ed_findings", "") not in ("", ss.get("last_suggested", "")):
        ss["findings_is_auto"] = False
    if ss.get("findings_is_auto", True):
        ss["ed_findings"] = suggested
    ss["last_suggested"] = suggested
    # ---------- Load saved edits ----------
    with st.expander("📂 Load saved edits (so you don't have to edit again)"):
        saved = st.file_uploader("Choose a saved edits file (.json)", type=["json"], key="edits_file")
        if saved is not None and ss.get("edits_loaded_id") != saved.file_id:
            try:
                e = json.loads(saved.getvalue().decode("utf-8"))
                for k in ("title", "role", "author", "findings", "recs"):
                    if k in e:
                        ss[f"ed_{k}"] = e[k]
                if e.get("sections"):
                    ss["ed_sections"] = [x for x in e["sections"] if x in SECTION_LABELS]
                ss["q_settings"] = e.get("questions", {})
                ss["edits_loaded_id"] = saved.file_id
                ss["findings_is_auto"] = False
                ss["q_editor_version"] = ss.get("q_editor_version", 0) + 1
                st.success("Edits loaded.")
            except Exception:
                st.error("This file could not be read. Please choose an edits file saved from this app.")

    # ---------- 1. Cover ----------
    with st.container(key="card_cover"):
        card_title("1 · Cover page")
        t1, t2, t3 = st.columns([2, 1.3, 1.3])
        t1.text_input("Report title", key="ed_title")
        t2.text_input("Role / school", key="ed_role")
        t3.text_input("Prepared by", key="ed_author")
    st.write("")

    # ---------- 2. Findings & recommendations ----------
    with st.container(key="card_notes"):
        card_title("2 · Key findings & recommendations")
        st.caption("Write one point per line. They appear on their own page in the PDF, PowerPoint and "
                   "Interactive report.")
        n1, n2 = st.columns(2)
        n1.text_area("Key findings", key="ed_findings", height=220,
                     placeholder="e.g. 47% of students do not get 8–10 hours of sleep, mostly due to study load")
        n2.text_area("Recommendations", key="ed_recs", height=220,
                     placeholder="e.g. Review hostel study and wake-up times to protect sleep")
        if n1.button("↺ Refill findings from the data"):
            ss["refill_findings"] = True
            st.rerun()
        if ss.get("findings_is_auto", True):
            n1.caption("✨ Drafted automatically from your data — edit or add to it.")
    st.write("")

    # ---------- 3. Questions ----------
    with st.container(key="card_questions"):
        card_title("3 · Questions in the report")
        st.caption("Untick **Show** to leave a question out, change its **Wording in report**, "
                   "and tick **Red bar** for key indicators.")
        qs = ss["q_settings"]
        base = pd.DataFrame({
            "Show": [qs.get(q, {}).get("show", True) for q in summary["Question"]],
            "Wording in report": [qs.get(q, {}).get("label", clean_question(q)) for q in summary["Question"]],
            "Red bar": [qs.get(q, {}).get("red", bool(h)) for q, h in zip(summary["Question"], summary["Highlight"])],
            "% Yes": summary["% Yes"].values,
            "Question in survey": summary["Question"].values,
        })
        edited = st.data_editor(
            base, hide_index=True, use_container_width=True, num_rows="fixed",
            key=f"q_editor_{report_name}_{ss.get('q_editor_version', 0)}",
            disabled=["% Yes", "Question in survey"],
            column_config={
                "Show": st.column_config.CheckboxColumn(width="small"),
                "Wording in report": st.column_config.TextColumn(width="large"),
                "Red bar": st.column_config.CheckboxColumn(width="small"),
                "% Yes": st.column_config.NumberColumn(format="%.0f%%", width="small"),
                "Question in survey": st.column_config.TextColumn(width="medium"),
            })
        for _, r in edited.iterrows():
            qs[r["Question in survey"]] = {"show": bool(r["Show"]), "label": str(r["Wording in report"]).strip(),
                                           "red": bool(r["Red bar"])}
    st.write("")

    # ---------- 4. Pages + download ----------
    with st.container(key="card_export"):
        card_title("4 · Pages and download")
        st.multiselect("Pages to include", list(SECTION_LABELS), key="ed_sections",
                       format_func=lambda k: SECTION_LABELS[k])

        # Build the edited report data
        summary_out = summary.copy()
        summary_out["Original"] = summary_out["Question"]
        summary_out["Highlight"] = [qs[q]["red"] for q in summary_out["Original"]]
        summary_out["Question"] = [qs[q]["label"] or clean_question(q) for q in summary_out["Original"]]
        summary_out = summary_out[[qs[q]["show"] for q in summary_out["Original"]]].reset_index(drop=True)
        notes = {"findings": lines_of(ss["ed_findings"]), "recommendations": lines_of(ss["ed_recs"])}
        sections = ss["ed_sections"] or list(SECTION_LABELS)
        pdf_title = "\n".join([ss["ed_title"].strip(), ss["ed_role"].strip(), ss["ed_author"].strip()])
        file_stem = (" - ".join(x.strip() for x in (ss["ed_title"], ss["ed_role"]) if x.strip()) or "Survey Report")
        file_stem = file_stem.replace("/", "-").replace("&", "and")

        edits = {"title": ss["ed_title"], "role": ss["ed_role"], "author": ss["ed_author"],
                 "findings": ss["ed_findings"], "recs": ss["ed_recs"], "sections": sections, "questions": qs}
        report_kwargs = dict(notes=notes, sections=sections)
        report_args = (df, pdf_title, summary_out, cat_cols, filters_text, group, chosen)

        # Forget reports made before the latest edits / sheet / filter change
        signature = (file.name, report_name, filters_text, group, tuple(chosen or []),
                     json.dumps(edits, sort_keys=True))
        if ss.get("report_signature") != signature:
            for ext in ("pdf", "pptx", "html"):
                ss.pop(ext, None)
            ss["report_signature"] = signature

        if "notes" in sections and not (notes["findings"] or notes["recommendations"]):
            st.warning("The **Key findings & recommendations** page will be left out because both boxes in "
                       "section 2 are empty.")
        elif "notes" not in sections:
            st.info("The Key findings & recommendations page is switched off in **Pages to include**.")
        st.caption("The Interactive report is a web page that works on phones and computers, even offline: "
                   "tap any chart to see the exact numbers. No names or written answers are included.")
        EXPORTS = [
            ("PDF", build_pdf, "pdf", "application/pdf"),
            ("PowerPoint", build_pptx, "pptx",
             "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
            ("Interactive report", build_html, "html", "text/html"),
        ]
        for col, (label, builder, ext, mime) in zip(st.columns(len(EXPORTS)), EXPORTS):
            with col:
                if st.button(f"Create {label}", use_container_width=True):
                    with st.spinner(f"Building {label}..."):
                        ss[ext] = builder(*report_args, **report_kwargs)
                if ext in ss:
                    st.download_button(f"⬇️ Download {label}", ss[ext], f"{file_stem}.{ext}",
                                       mime, use_container_width=True)

        st.divider()
        c1, c2 = st.columns(2)
        c1.download_button("💾 Save my edits (to reuse next time)", json.dumps(edits, indent=2, ensure_ascii=False),
                           f"{file_stem} - edits.json", "application/json", use_container_width=True)
        c2.download_button("Download Yes/No summary (CSV)",
                           summary_out.drop(columns=["Highlight", "Original"]).to_csv(index=False), "summary.csv",
                           use_container_width=True)
