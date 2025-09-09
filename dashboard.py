import os
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px

# -----------------------------
# BRAND
# -----------------------------
COLOR_DARK   = "#1f2951"  # mørk primær
COLOR_ACCENT = "#d6a550"  # detalje
COLOR_BLUE   = "#004899"  # detalje
COLOR_WHITE  = "#ffffff"

PLOTLY_COLORWAY = [COLOR_ACCENT, COLOR_BLUE, COLOR_DARK, "#888888", "#bbbbbb"]

# -----------------------------
# STREAMLIT CONFIG
# -----------------------------
st.set_page_config(page_title="Vejlednings-dashboard", page_icon="📊", layout="wide")

# Branded header (uden logo)
st.markdown(
    f"""
    <div style="background:{COLOR_DARK}; padding:14px 18px; border-radius:12px; margin-bottom:14px;">
        <h1 style="color:{COLOR_WHITE}; margin:0; font-weight:700; font-size:22px;">📊 Vejlednings-dashboard</h1>
        <p style="color:{COLOR_WHITE}; opacity:.85; margin:4px 0 0 0; font-size:13px;">ErhvervsForum Holstebro · live overblik</p>
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------
# CONSTANTS
# -----------------------------
DATE_COL_CANDIDATES = ["Mødedato", "Mødedato_Møder", "Dato"]
CREATOR_COL = "Oprettet_af_Møder"
COMPANY_COL = "Firmanavn_Virksomheder"
TOPIC_COL   = "Emner_Møder"
TITLE_COL   = "Titel_Møder"
KOMMUNE_COL = "Kommune_Virksomheder"

CORE_NAMES = ["Victor", "Jan", "Mette", "Kristina", "Sara", "Peter"]
BUCKET_OTHER = "Erhvervshus Midtjylland"

# -----------------------------
# HELPERS
# -----------------------------
@st.cache_data(show_spinner=False)
def load_excel_merge_two_sheets(path: str = None, uploaded=None) -> pd.DataFrame:
    """
    Læs både ark 1 og 2, hvis 2 findes. Ark 2 (EHM) tildeles Vejleder = Erhvervshus Midtjylland.
    """
    if uploaded is not None:
        dfs = pd.read_excel(uploaded, sheet_name=None)
    elif path and os.path.exists(path):
        dfs = pd.read_excel(path, sheet_name=None)
    else:
        return pd.DataFrame()

    if not isinstance(dfs, dict) or len(dfs) == 0:
        return pd.DataFrame()

    sheet_names = list(dfs.keys())
    df1 = dfs[sheet_names[0]].copy()
    df1["__sheet__"] = sheet_names[0]
    merged = df1

    if len(sheet_names) >= 2:
        df2 = dfs[sheet_names[1]].copy()
        df2["__sheet__"] = sheet_names[1]

        commons = [c for c in df2.columns if c in df1.columns]
        if commons:
            df2 = df2[commons + ["__sheet__"]]

        df2["Vejleder"] = BUCKET_OTHER
        merged = pd.concat([merged, df2], ignore_index=True, sort=False)

    # Map kerne-navne for rækker fra ark 1
    if CREATOR_COL in merged.columns:
        def clean_vejleder(val):
            if pd.isna(val):
                return BUCKET_OTHER
            s = str(val).lower()
            for name in CORE_NAMES:
                if name.lower() in s:
                    return name
            return BUCKET_OTHER

        if "Vejleder" not in merged.columns:
            merged["Vejleder"] = merged[CREATOR_COL].apply(clean_vejleder)
        else:
            mask_empty = merged["Vejleder"].isna() | (merged["Vejleder"] == "")
            merged.loc[mask_empty, "Vejleder"] = merged.loc[mask_empty, CREATOR_COL].apply(clean_vejleder)

    return merged

def pick_date_col(df: pd.DataFrame) -> str | None:
    for c in DATE_COL_CANDIDATES:
        if c in df.columns:
            return c
    return None

def parse_dates(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.dropna(subset=[date_col])
        df["Dato"] = df[date_col]
        df["Måned"] = df["Dato"].apply(lambda d: pd.Timestamp(year=d.year, month=d.month, day=1) if pd.notnull(d) else pd.NaT)
    return df

def fmt_int(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except Exception:
        return "—"

def plot_line(df, x, y, color=None):
    fig = px.line(df, x=x, y=y, color=color, markers=True, color_discrete_sequence=PLOTLY_COLORWAY)
    fig.update_layout(margin=dict(l=0, r=0, t=6, b=6), height=380, plot_bgcolor="rgba(0,0,0,0)")
    return fig

def plot_bar(df, x, y, color=None, text=None):
    fig = px.bar(df, x=x, y=y, color=color, text=text, color_discrete_sequence=PLOTLY_COLORWAY)
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(margin=dict(l=0, r=0, t=6, b=6), height=320, plot_bgcolor="rgba(0,0,0,0)")
    return fig

# -----------------------------
# SIDEBAR: DATA
# -----------------------------
st.sidebar.header("🔧 Datakilde")
default_path = st.sidebar.text_input("Sti til Excel-fil", value=os.environ.get("DASHBOARD_DATA_PATH", ""))
uploaded_file = st.sidebar.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
auto_refresh = st.sidebar.number_input("Auto-refresh (sek.)", min_value=0, max_value=3600, value=0, step=5)
if auto_refresh > 0:
    st.sidebar.caption("Browseren kan også auto-reloade via skærmsoftware.")
    st.experimental_set_query_params(_=datetime.now().isoformat())

df = load_excel_merge_two_sheets(default_path, uploaded_file)
if df.empty:
    st.warning("Upload en Excel-fil eller angiv en sti i sidebaren. (Denne version læser ark 1 + ark 2 og bucket’er EHM)")
    st.stop()

date_col = pick_date_col(df)
df = parse_dates(df, date_col)

# -----------------------------
# FILTRE
# -----------------------------
st.sidebar.header("🔎 Filtre")
if "Vejleder" not in df.columns:
    df["Vejleder"] = BUCKET_OTHER

unique_vejledere = sorted(df["Vejleder"].dropna().unique(), key=lambda x: (x != BUCKET_OTHER, x))
selected_vejledere = st.sidebar.multiselect("Vejleder(e)", options=list(unique_vejledere), default=unique_vejledere)

if selected_vejledere:
    df = df[df["Vejleder"].isin(selected_vejledere)]

if "Dato" in df.columns and not df["Dato"].isna().all():
    min_dato = pd.to_datetime(df["Dato"].min()).date()
    max_dato = pd.to_datetime(df["Dato"].max()).date()
    date_range = st.sidebar.date_input("Datointerval", value=(min_dato, max_dato))
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        mask = (df["Dato"].dt.date >= start_date) & (df["Dato"].dt.date <= end_date)
        df = df[mask]

# -----------------------------
# LAYOUT-KONTROL
# -----------------------------
st.sidebar.header("🧩 Layout")
SECTIONS = ["KPIs", "Udvikling", "Fordelinger", "Seneste 10"]
if "section_order" not in st.session_state:
    st.session_state.section_order = SECTIONS.copy()

opt1 = st.sidebar.selectbox("1. sektion", SECTIONS, index=0, key="s1")
rem2 = [s for s in SECTIONS if s != opt1]
opt2 = st.sidebar.selectbox("2. sektion", rem2, index=0, key="s2")
rem3 = [s for s in rem2 if s != opt2]
opt3 = st.sidebar.selectbox("3. sektion", rem3, index=0, key="s3")
rem4 = [s for s in rem3 if s != opt3]
opt4 = st.sidebar.selectbox("4. sektion", rem4, index=0, key="s4")
st.session_state.section_order = [opt1, opt2, opt3, opt4]

# -----------------------------
# SECTIONS
# -----------------------------
def render_kpis(df: pd.DataFrame):
    st.subheader("Nøgletal")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("📈 Total (filteret)", fmt_int(len(df)))
    with c2:
        if COMPANY_COL in df.columns:
            st.metric("🏷️ Unikke virksomheder", fmt_int(df[COMPANY_COL].nunique()))
    with c3:
        if "Dato" in df.columns:
            sidste_30 = (df["Dato"] >= (pd.Timestamp.now() - pd.Timedelta(days=30))).sum()
            st.metric("🗓️ Seneste 30 dage", fmt_int(sidste_30))
    with c4:
        if "Dato" in df.columns and not df.empty:
            row = df.sort_values("Dato", ascending=False).iloc[0]
            txt = row["Dato"].strftime("%Y-%m-%d")
            if COMPANY_COL in df.columns and pd.notna(row.get(COMPANY_COL, None)):
                txt += f" · {row[COMPANY_COL]}"
            st.metric("⏱️ Seneste vejledning", txt)
    with c5:
        ehm = (df["Vejleder"] == BUCKET_OTHER).sum()
        st.metric("🏢 EHM (filteret)", fmt_int(ehm))

def render_trend(df: pd.DataFrame):
    st.subheader("Udvikling over tid")
    if "Måned" in df.columns and not df["Måned"].isna().all():
        trend = (
            df.dropna(subset=["Måned"])
              .groupby(["Måned", "Vejleder"], dropna=False)
              .size()
              .reset_index(name="Antal")
              .sort_values("Måned")
        )
        fig = plot_line(trend, x="Måned", y="Antal", color="Vejleder")
        st.plotly_chart(fig, use_container_width=True)

def render_distributions(df: pd.DataFrame):
    st.subheader("Fordelinger")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**Vejledninger pr. vejleder**")
        by_vej = df.groupby("Vejleder", dropna=False).size().reset_index(name="Antal").sort_values("Antal", ascending=False)
        fig = plot_bar(by_vej, x="Vejleder", y="Antal", text="Antal")
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        if TOPIC_COL in df.columns:
            st.markdown("**Top emner/temaer**")
            by_topic = (df.assign(**{TOPIC_COL: df[TOPIC_COL].fillna("Ukendt")})
                          .groupby(TOPIC_COL, dropna=False).size()
                          .reset_index(name="Antal")
                          .sort_values("Antal", ascending=False).head(12))
            fig = plot_bar(by_topic, x=TOPIC_COL, y="Antal", text="Antal")
            fig.update_layout(xaxis_tickangle=-30)
            st.plotly_chart(fig, use_container_width=True)
    with col3:
        if KOMMUNE_COL in df.columns:
            st.markdown("**Kommunefordeling**")
            by_kom = (df.assign(**{KOMMUNE_COL: df[KOMMUNE_COL].fillna("Ukendt")})
                        .groupby(KOMMUNE_COL, dropna=False).size()
                        .reset_index(name="Antal")
                        .sort_values("Antal", ascending=False).head(12))
            fig = plot_bar(by_kom, x=KOMMUNE_COL, y="Antal", text="Antal")
            fig.update_layout(xaxis_tickangle=-30)
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
    tbl = tbl.rename(columns={COMPANY_COL:"Virksomhed", TOPIC_COL:"Emne", TITLE_COL:"Titel"})
    st.dataframe(tbl, use_container_width=True)

# -----------------------------
# RENDER ORDER
# -----------------------------
for sec in st.session_state.section_order:
    if sec == "KPIs":
        render_kpis(df)
    elif sec == "Udvikling":
        render_trend(df)
    elif sec == "Fordelinger":
        render_distributions(df)
    elif sec == "Seneste 10":
        render_latest(df)

st.caption("Farvetema: #1f2951 · #d6a550 · #004899 · Data: ark 1 + ark 2 (EHM).")
