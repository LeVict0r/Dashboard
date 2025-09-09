
import os
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Vejlednings-dashboard", page_icon="📊", layout="wide")

DATE_COL_CANDIDATES = ["Mødedato", "Mødedato_Møder"]
CREATOR_COL = "Oprettet_af_Møder"
COMPANY_COL = "Firmanavn_Virksomheder"
TOPIC_COL = "Emner_Møder"
TITLE_COL = "Titel_Møder"
KOMMUNE_COL = "Kommune_Virksomheder"

CORE_NAMES = ["Victor", "Jan", "Mette", "Kristina", "Sara", "Peter"]
BUCKET_OTHER = "Erhvervshus Midtjylland"

@st.cache_data(show_spinner=False)
def load_excel(path: str = None, uploaded=None) -> pd.DataFrame:
    if uploaded is not None:
        df = pd.read_excel(uploaded)
    elif path and os.path.exists(path):
        df = pd.read_excel(path)
    else:
        return pd.DataFrame()
    df.columns = [str(c) for c in df.columns]
    return df

def pick_date_col(df: pd.DataFrame) -> str | None:
    for c in DATE_COL_CANDIDATES:
        if c in df.columns:
            return c
    return None

def parse_dates(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    return df

def clean_vejleder(val: str | None) -> str:
    if not isinstance(val, str) or val.strip() == "":
        return BUCKET_OTHER
    lower = val.lower()
    for name in CORE_NAMES:
        if name.lower() in lower:
            return name
    if "erhvervshus" in lower and "midt" in lower:
        return BUCKET_OTHER
    return BUCKET_OTHER

def month_floor(dt: pd.Timestamp) -> pd.Timestamp:
    return pd.Timestamp(year=dt.year, month=dt.month, day=1)

def kpi_metric(label: str, value, help_text: str | None = None):
    st.metric(label, value, help=help_text)

def format_int(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except Exception:
        return "—"

st.sidebar.header("🔧 Datakilde")
default_path = st.sidebar.text_input("Sti til Excel-fil", value=os.environ.get("DASHBOARD_DATA_PATH", ""))
uploaded_file = st.sidebar.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

df = load_excel(default_path, uploaded_file)
if df.empty:
    st.warning("Upload en Excel-fil eller angiv en sti i sidebaren.")
    st.stop()

date_col = pick_date_col(df)
df = parse_dates(df, date_col)
if date_col:
    df = df.dropna(subset=[date_col])

if CREATOR_COL in df.columns:
    df["Vejleder"] = df[CREATOR_COL].apply(clean_vejleder)
else:
    df["Vejleder"] = BUCKET_OTHER

if date_col:
    df["Dato"] = df[date_col]
    df["Måned"] = df["Dato"].apply(lambda d: month_floor(d) if pd.notnull(d) else pd.NaT)

st.sidebar.header("🔎 Filtre")
unique_vejledere = sorted(df["Vejleder"].dropna().unique())
selected_vejledere = st.sidebar.multiselect("Vejleder(e)", options=list(unique_vejledere), default=unique_vejledere)

if selected_vejledere:
    df = df[df["Vejleder"].isin(selected_vejledere)]

st.sidebar.header("🧩 Layout")
SECTIONS = ["KPIs", "Udvikling", "Fordelinger", "Seneste 10"]
if "section_order" not in st.session_state:
    st.session_state.section_order = SECTIONS.copy()

opt1 = st.sidebar.selectbox("1. sektion", SECTIONS, index=0)
remaining2 = [s for s in SECTIONS if s != opt1]
opt2 = st.sidebar.selectbox("2. sektion", remaining2, index=0)
remaining3 = [s for s in remaining2 if s != opt2]
opt3 = st.sidebar.selectbox("3. sektion", remaining3, index=0)
remaining4 = [s for s in remaining3 if s != opt3]
opt4 = st.sidebar.selectbox("4. sektion", remaining4, index=0)
st.session_state.section_order = [opt1, opt2, opt3, opt4]

def render_kpis(df: pd.DataFrame):
    st.subheader("Nøgletal")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        kpi_metric("📈 Total vejledninger", format_int(len(df)))
    with c2:
        if COMPANY_COL in df.columns:
            kpi_metric("🏷️ Unikke virksomheder", format_int(df[COMPANY_COL].nunique()))
    with c3:
        if "Dato" in df.columns:
            sidste_30 = (df["Dato"] >= (pd.Timestamp.now() - pd.Timedelta(days=30))).sum()
            kpi_metric("🗓️ Seneste 30 dage", format_int(sidste_30))
    with c4:
        if "Dato" in df.columns and not df.empty:
            seneste_row = df.sort_values("Dato", ascending=False).iloc[0]
            seneste_txt = seneste_row["Dato"].strftime("%Y-%m-%d")
            if COMPANY_COL in df.columns and pd.notna(seneste_row.get(COMPANY_COL, None)):
                seneste_txt += f" · {seneste_row[COMPANY_COL]}"
            kpi_metric("⏱️ Seneste vejledning", seneste_txt)
    with c5:
        total_eh = (df["Vejleder"] == BUCKET_OTHER).sum()
        kpi_metric("🏢 Erhvervshus Midtjylland", format_int(total_eh))

def render_trend(df: pd.DataFrame):
    st.subheader("Udvikling over tid")
    if "Måned" in df.columns and not df["Måned"].isna().all():
        trend = (
            df.dropna(subset=["Måned"])
              .groupby(["Måned", "Vejleder"])
              .size()
              .reset_index(name="Antal")
        )
        fig = px.line(trend, x="Måned", y="Antal", color="Vejleder", markers=True)
        st.plotly_chart(fig, use_container_width=True)

def render_distributions(df: pd.DataFrame):
    st.subheader("Fordelinger")
    col1, col2, col3 = st.columns(3)
    with col1:
        by_vej = df.groupby("Vejleder").size().reset_index(name="Antal").sort_values("Antal", ascending=False)
        fig = px.bar(by_vej, x="Vejleder", y="Antal", text="Antal")
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        if TOPIC_COL in df.columns:
            by_topic = df.groupby(TOPIC_COL).size().reset_index(name="Antal").sort_values("Antal", ascending=False).head(10)
            fig = px.bar(by_topic, x=TOPIC_COL, y="Antal", text="Antal")
            st.plotly_chart(fig, use_container_width=True)
    with col3:
        if KOMMUNE_COL in df.columns:
            by_kom = df.groupby(KOMMUNE_COL).size().reset_index(name="Antal").sort_values("Antal", ascending=False).head(10)
            fig = px.bar(by_kom, x=KOMMUNE_COL, y="Antal", text="Antal")
            st.plotly_chart(fig, use_container_width=True)

def render_latest(df: pd.DataFrame):
    st.subheader("Seneste 10 vejledninger")
    cols = ["Dato", "Vejleder"]
    if COMPANY_COL in df.columns: cols.append(COMPANY_COL)
    if TOPIC_COL in df.columns: cols.append(TOPIC_COL)
    if TITLE_COL in df.columns: cols.append(TITLE_COL)
    tbl = df.sort_values("Dato", ascending=False)[cols].head(10)
    if "Dato" in tbl.columns:
        tbl["Dato"] = tbl["Dato"].dt.strftime("%Y-%m-%d")
    st.dataframe(tbl, use_container_width=True)

st.title("📊 Vejlednings-dashboard")

for sec in st.session_state.section_order:
    if sec == "KPIs":
        render_kpis(df)
    elif sec == "Udvikling":
        render_trend(df)
    elif sec == "Fordelinger":
        render_distributions(df)
    elif sec == "Seneste 10":
        render_latest(df)

st.caption("Layout kan tilpasses i sidebaren. KPI viser også total for Erhvervshus Midtjylland.")
