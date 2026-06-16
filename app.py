import streamlit as st
import subprocess
import os
import sys
import glob
import base64
import pandas as pd

st.set_page_config(
    page_title="Jumbo Marketing Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Logo helper ───────────────────────────────────────────────────────────────
def _logo_b64():
    path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

_logo = _logo_b64()
_logo_html = (
    f'<img src="data:image/png;base64,{_logo}" '
    f'style="height:56px;border-radius:12px;flex-shrink:0;'
    f'background:#fff;padding:6px;box-shadow:0 0 0 1.5px rgba(255,255,255,0.15);">'
    if _logo else ""
)

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
[data-testid="stDecoration"] {display: none;}
[data-testid="stDeployButton"] {display: none;}
header {background: transparent !important; border: none !important; box-shadow: none !important;}

section[data-testid="stSidebar"] {
    background: #11131f;
    border-right: 1px solid #1d1f35;
}

div[data-testid="metric-container"] {
    background: #161828;
    border: 1px solid #252845;
    border-radius: 12px;
    padding: 1rem 1.2rem;
    box-shadow: 0 4px 16px rgba(0,0,0,0.35);
}
div[data-testid="metric-container"] label {
    color: #6b7190 !important;
    font-size: 0.72rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #e8eaf6 !important;
    font-size: 1.7rem !important;
    font-weight: 700 !important;
}

.hero {
    background: linear-gradient(135deg, #1a053a 0%, #0d1547 55%, #091a3d 100%);
    border-radius: 14px;
    padding: 1.8rem 2.2rem;
    border: 1px solid #2a1f4a;
    margin-bottom: 1.5rem;
}
.hero h1 {
    font-size: 2rem;
    font-weight: 800;
    margin: 0 0 0.05rem 0;
    background: linear-gradient(90deg, #ffffff 0%, #c4b5fd 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    letter-spacing: -0.02em;
}
.hero-byline {
    color: #7880a8;
    font-size: 0.85rem;
    font-weight: 400;
    margin: 0 0 0.9rem 0;
}
.pills { display: flex; gap: 0.4rem; flex-wrap: wrap; }
.pill {
    padding: 0.28rem 0.85rem;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.02em;
}
.pill-active {
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    color: #fff;
    box-shadow: 0 0 12px rgba(99,102,241,0.45);
}
.pill-soon {
    background: rgba(255,255,255,0.05);
    border: 1px solid #3a3f6a;
    color: #7880a8;
}

.section-label {
    font-size: 0.7rem;
    font-weight: 600;
    color: #454870;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin: 1.4rem 0 0.6rem 0;
}

.top-cat-box {
    background: #161828;
    border-left: 3px solid #6366f1;
    border-radius: 8px;
    padding: 0.7rem 1rem;
    color: #a5b4fc;
    font-size: 0.88rem;
    font-weight: 500;
    margin: 0.8rem 0;
}

.empty-state {
    text-align: center;
    padding: 3rem;
    color: #3b4070;
    border: 1.5px dashed #1d1f35;
    border-radius: 12px;
    background: #11131f;
}
.empty-state .icon { font-size: 2.5rem; }
.empty-state p { margin: 0.5rem 0 0 0; font-size: 0.9rem; }
</style>""", unsafe_allow_html=True)

# ── Hero Header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div>
        <h1>Jumbo Marketing Analytics</h1>
        <p class="hero-byline">by Somya Mishra</p>
    </div>
    <div class="pills">
        <span class="pill pill-active">🔔 Push Notifications</span>
        <span class="pill pill-soon">📱 In-App &nbsp;· Soon</span>
        <span class="pill pill-soon">🗺️ Journeys &nbsp;· Soon</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    if _logo:
        st.image("logo.png", width=60)
    st.markdown("### 📋 How to Use")
    st.markdown("""
1. Upload your raw campaign file (`.xlsx` or `.csv`)
2. Click **Run Automation**
3. Review summary metrics
4. Download report as **XLSX** or **CSV**
    """)
    st.divider()
    with st.expander("Required Columns"):
        st.markdown("""
- Campaign Name / ID
- Channel, Variant, Device
- Title, Message
- Start Date / Time, Run Date
- Total Sent / Viewed / Clicked / Delivered
- Click through conversions
- Control group metrics
- Influenced Conversions / Revenue
        """)
    st.divider()
    st.caption("Push Notifications · v1.0")

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-label">Upload</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload",
    type=["xlsx", "csv"],
    label_visibility="collapsed"
)

if not uploaded_file:
    st.markdown("""
        <div class="empty-state">
            <div class="icon">📂</div>
            <p>Upload a campaign file (.xlsx or .csv) to get started</p>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

st.markdown(f"""
    <div style="display:flex;align-items:center;gap:0.75rem;
                background:#0e1f14;border:1px solid #1a3a22;
                border-radius:8px;padding:0.6rem 1rem;margin-bottom:0.5rem">
        <span style="color:#4ade80;font-size:1rem">✅</span>
        <span style="color:#86efac;font-size:0.9rem;font-weight:500">{uploaded_file.name}</span>
    </div>
""", unsafe_allow_html=True)
run = st.button("▶  Run Automation", use_container_width=True, type="primary")

uploaded_path = os.path.basename(uploaded_file.name)
with open(uploaded_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

stem = os.path.splitext(uploaded_path)[0]
output_file = stem + "_Analysis.xlsx"
csv_file    = stem + "_Analysis.csv"

if not run:
    st.stop()

# ── Cleanup old outputs ───────────────────────────────────────────────────────
for pattern in ("*_Analysis.xlsx", "*_Analysis.csv"):
    for file in glob.glob(pattern):
        try:
            os.remove(file)
        except OSError:
            pass

# ── Run pn.py ─────────────────────────────────────────────────────────────────
with st.spinner("Running analysis…"):
    result = subprocess.run(
        [sys.executable, "pn.py", uploaded_path],
        capture_output=True,
        text=True
    )

if result.returncode != 0:
    st.error("❌  Automation Failed")
    lines = result.stderr.strip().splitlines()
    error_start = 0
    for i, line in enumerate(lines):
        if line and not line.startswith(" ") and not line.startswith("Traceback"):
            error_start = i
    error_msg = "\n".join(lines[error_start:]).strip() if lines else result.stderr.strip()
    st.warning(error_msg)
    st.stop()

if result.stdout:
    with st.expander("Logs", expanded=False):
        st.code(result.stdout)

if not os.path.exists(output_file):
    st.error(f"Output file not found: {output_file}")
    st.stop()

# ── Summary Metrics ───────────────────────────────────────────────────────────
st.markdown('<p class="section-label">Campaign Summary</p>', unsafe_allow_html=True)

if os.path.exists(csv_file):
    df = pd.read_csv(csv_file)

    total_campaigns = df["Campaign ID"].nunique()        if "Campaign ID"       in df.columns else 0
    total_sent      = int(df["Total Sent(users)"].sum()) if "Total Sent(users)" in df.columns else 0
    total_converted = int(df["Total Converted"].sum())   if "Total Converted"   in df.columns else 0
    avg_ctr         = df["CTR%"].mean()                  if "CTR%"              in df.columns else 0
    avg_lift        = df["Relative Lift"].mean()         if "Relative Lift"     in df.columns else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Campaigns",         f"{total_campaigns:,}")
    m2.metric("Total Sent",        f"{total_sent:,}")
    m3.metric("Total Converted",   f"{total_converted:,}")
    m4.metric("Avg CTR",           f"{avg_ctr:.2f}%")
    m5.metric("Avg Relative Lift", f"{avg_lift:.1f}%")

    if "Category" in df.columns and "Total Converted" in df.columns:
        top_cat = (
            df.groupby("Category")["Total Converted"]
            .sum().idxmax()
        )
        st.markdown(
            f'<div class="top-cat-box">🏆 Top Category by Conversions — <strong>{top_cat}</strong></div>',
            unsafe_allow_html=True
        )

    # ── Top Campaigns Table ───────────────────────────────────────────────────
    table_cols = [
        "Campaign Name", "Variant", "Device",
        "CTR%", "TC%", "Total Converted",
        "Relative Lift", "Incremental Conversions"
    ]
    available = [c for c in table_cols if c in df.columns]

    if available:
        st.markdown('<p class="section-label">Top 10 Campaigns by Conversions</p>', unsafe_allow_html=True)
        sort_col = "Total Converted" if "Total Converted" in df.columns else available[0]
        top_df = (
            df[available]
            .sort_values(sort_col, ascending=False)
            .head(10)
            .reset_index(drop=True)
        )
        top_df.index += 1

        fmt = {}
        if "CTR%"                    in top_df.columns: fmt["CTR%"]                    = "{:.2f}%"
        if "TC%"                     in top_df.columns: fmt["TC%"]                     = "{:.2f}%"
        if "Relative Lift"           in top_df.columns: fmt["Relative Lift"]           = "{:.1f}%"
        if "Incremental Conversions" in top_df.columns: fmt["Incremental Conversions"] = "{:.1f}"
        if "Total Converted"         in top_df.columns: fmt["Total Converted"]         = "{:,.0f}"

        st.dataframe(
            top_df.style.format(fmt),
            use_container_width=True
        )

# ── Download ──────────────────────────────────────────────────────────────────
st.divider()
st.markdown('<p class="section-label">Download Report</p>', unsafe_allow_html=True)

dl1, dl2 = st.columns(2)

with dl1:
    with open(output_file, "rb") as f:
        xlsx_bytes = f.read()
    st.download_button(
        label="📥  Download as XLSX",
        data=xlsx_bytes,
        file_name=os.path.basename(output_file),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

with dl2:
    if os.path.exists(csv_file):
        with open(csv_file, "rb") as f:
            csv_bytes = f.read()
        st.download_button(
            label="📄  Download as CSV",
            data=csv_bytes,
            file_name=os.path.basename(csv_file),
            mime="text/csv",
            use_container_width=True
        )
