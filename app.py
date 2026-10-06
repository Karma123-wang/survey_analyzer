"""
Survey Analyser - upload a Google Forms / Excel survey export and get instant analysis.
Run:  pip install streamlit pandas openpyxl plotly
      streamlit run app.py
"""
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Survey Analyser", layout="wide")
st.title("Survey Analyser")

# Columns that identify a person - never shown in results
PRIVATE_EXACT = {"name", "name:", "timestamp", "si.no", "sl.no", "cid", "phone", "email"}
PRIVATE_WORDS = ["address", "please write the name"]


def is_private(col: str) -> bool:
    c = col.lower().strip()
    return c in PRIVATE_EXACT or any(w in c for w in PRIVATE_WORDS)


def is_yes_no(series: pd.Series) -> bool:
    vals = set(series.dropna().astype(str).str.strip().unique())
    return 0 < len(vals) <= 2 and vals <= {"Yes", "No"}


def is_category(series: pd.Series) -> bool:
    return 1 < series.nunique() <= 12 and not is_yes_no(series)


# ---------- 1. Upload ----------
file = st.file_uploader("Upload survey file (.xlsx or .csv)", type=["xlsx", "csv"])
if not file:
    st.info("Upload a file to begin.")
    st.stop()

if file.name.endswith(".csv"):
    df = pd.read_csv(file)
else:
    xls = pd.ExcelFile(file)
    sheet = st.selectbox("Sheet", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet)

df = df.dropna(how="all")
df.columns = [str(c).strip() for c in df.columns]
for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.strip().replace({"nan": None})

usable = [c for c in df.columns if not is_private(c)]
yn_cols = [c for c in usable if is_yes_no(df[c])]
cat_cols = [c for c in usable if is_category(df[c])]

# ---------- 2. Filters ----------
st.sidebar.header("Filters")
for c in cat_cols[:6]:
    opts = sorted(df[c].dropna().unique())
    pick = st.sidebar.multiselect(c, opts, default=opts)
    df = df[df[c].isin(pick)]

st.metric("Responses", len(df))
if df.empty:
    st.warning("No responses match the filters.")
    st.stop()

# ---------- 3. Profile of respondents ----------
st.header("Respondent profile")
cols = st.columns(3)
for i, c in enumerate(cat_cols):
    counts = df[c].value_counts().reset_index()
    counts.columns = [c, "count"]
    fig = px.pie(counts, names=c, values="count", title=c, hole=0.4)
    cols[i % 3].plotly_chart(fig, use_container_width=True)

# ---------- 4. Yes / No questions ----------
st.header("Yes / No questions")
summary = pd.DataFrame({
    "Question": yn_cols,
    "% Yes": [(df[c] == "Yes").mean() * 100 for c in yn_cols],
    "Yes": [(df[c] == "Yes").sum() for c in yn_cols],
    "No": [(df[c] == "No").sum() for c in yn_cols],
}).round(1)
fig = px.bar(summary, x="% Yes", y="Question", orientation="h", height=28 * len(yn_cols) + 100)
fig.update_layout(yaxis={"autorange": "reversed"})
st.plotly_chart(fig, use_container_width=True)
st.dataframe(summary, use_container_width=True, hide_index=True)

# ---------- 5. Compare groups ----------
st.header("Compare groups")
if cat_cols and yn_cols:
    group = st.selectbox("Break down by", cat_cols)
    chosen = st.multiselect("Questions", yn_cols, default=yn_cols[:5])
    if chosen:
        table = df.groupby(group)[chosen].apply(lambda g: (g == "Yes").mean() * 100).round(0)
        table["n"] = df.groupby(group).size()
        st.dataframe(table, use_container_width=True)
        long = table.drop(columns="n").reset_index().melt(id_vars=group, var_name="Question", value_name="% Yes")
        st.plotly_chart(px.bar(long, x="Question", y="% Yes", color=group, barmode="group"), use_container_width=True)

# ---------- 6. Free-text reasons ----------
st.header("Written reasons")
text_cols = [c for c in usable if c not in yn_cols and c not in cat_cols and df[c].nunique() > 12]
if text_cols:
    tc = st.selectbox("Column", text_cols)
    answers = df[tc].dropna()
    answers = answers[answers.str.len() > 3]  # skip "-", "no", etc.
    keyword = st.text_input("Count answers containing a keyword (e.g. study, hostel, exam)")
    if keyword:
        hits = answers[answers.str.lower().str.contains(keyword.lower())]
        st.write(f"**{len(hits)}** of {len(answers)} answers mention '{keyword}'")
        answers = hits
    st.dataframe(answers.reset_index(drop=True), use_container_width=True)

# ---------- 7. Download ----------
st.download_button("Download Yes/No summary (CSV)", summary.to_csv(index=False), "summary.csv")
