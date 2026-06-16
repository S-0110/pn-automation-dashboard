import streamlit as st
import subprocess
import os
import sys
import glob
import pandas as pd

st.set_page_config(
    page_title="Jumbo Marketing Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <div style='text-align:center; padding: 1.2rem 0 0.4rem 0'>
        <h1 style='font-size:2.2rem; margin-bottom:0'>📊 Jumbo Marketing Analytics Dashboard</h1>
        <p style='color:#888; margin-top:0.3rem; font-size:0.95rem'>Created by Somya Mishra</p>
    </div>
""", unsafe_allow_html=True)

st.divider()

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ℹ️ How to Use")
    st.markdown("""
1. Upload your raw campaign `.xlsx` file
2. Click **Run Automation**
3. Review the summary metrics
4. Download the full report
    """)
    st.divider()
    st.markdown("**Required Columns**")
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
    st.caption("Jumbo Marketing Analytics · Built with Streamlit")

# ── Upload ────────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload Raw Campaign File",
    type=["xlsx"],
    help="Export from your marketing platform"
)

if not uploaded_file:
    st.markdown("""
        <div style='text-align:center; padding:3rem; color:#888'>
            <div style='font-size:3rem'>📂</div>
            <p>Upload a campaign Excel file to get started</p>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

col_name, col_btn = st.columns([4, 1])
with col_name:
    st.success(f"✅ Uploaded: **{uploaded_file.name}**")
with col_btn:
    run = st.button("▶ Run Automation", use_container_width=True, type="primary")

uploaded_path = os.path.basename(uploaded_file.name)
with open(uploaded_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

if not run:
    st.stop()

# ── Cleanup old outputs ───────────────────────────────────────────────────────
for pattern in ("*_Analysis.xlsx", "*_Analysis.csv"):
    for file in glob.glob(pattern):
        try:
            os.remove(file)
        except OSError:
            pass

# ── Run script ────────────────────────────────────────────────────────────────
with st.spinner("Running analysis…"):
    result = subprocess.run(
        [sys.executable, "pn.py", uploaded_path],
        capture_output=True,
        text=True
    )

if result.returncode != 0:
    st.error("❌ Automation Failed")
    st.code(result.stderr)
    st.stop()

if result.stdout:
    with st.expander("Logs", expanded=False):
        st.code(result.stdout)

output_file = uploaded_path.replace(".xlsx", "_Analysis.xlsx")
csv_file    = uploaded_path.replace(".xlsx", "_Analysis.csv")

if not os.path.exists(output_file):
    st.error(f"Output file not found: {output_file}")
    st.stop()

st.success("✅ Report Generated Successfully")

# ── Metrics ───────────────────────────────────────────────────────────────────
if os.path.exists(csv_file):
    df = pd.read_csv(csv_file)

    st.markdown("### 📈 Campaign Summary")

    total_campaigns = df["Campaign ID"].nunique()           if "Campaign ID"                    in df.columns else 0
    total_sent      = int(df["Total Sent(users)"].sum())    if "Total Sent(users)"              in df.columns else 0
    total_converted = int(df["Total Converted"].sum())      if "Total Converted"                in df.columns else 0
    avg_ctr         = df["CTR%"].mean()                     if "CTR%"                           in df.columns else 0
    avg_lift        = df["Relative Lift"].mean()            if "Relative Lift"                  in df.columns else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Campaigns",       f"{total_campaigns:,}")
    m2.metric("Total Sent",      f"{total_sent:,}")
    m3.metric("Total Converted", f"{total_converted:,}")
    m4.metric("Avg CTR",         f"{avg_ctr:.2f}%")
    m5.metric("Avg Relative Lift", f"{avg_lift:.1f}%")

    if "Category" in df.columns and "Total Converted" in df.columns:
        top_category = (
            df.groupby("Category")["Total Converted"]
            .sum()
            .idxmax()
        )
        st.info(f"🏆 Top Category by Conversions: **{top_category}**")

    # ── Top Campaigns Table ───────────────────────────────────────────────────
    table_cols = [
        "Campaign Name", "Variant", "Device",
        "CTR%", "TC%", "Total Converted",
        "Relative Lift", "Incremental Conversions"
    ]
    available = [c for c in table_cols if c in df.columns]

    if available:
        st.markdown("### 🔝 Top 10 Campaigns by Conversions")
        sort_col = "Total Converted" if "Total Converted" in df.columns else available[0]
        top_df = (
            df[available]
            .sort_values(sort_col, ascending=False)
            .head(10)
            .reset_index(drop=True)
        )
        top_df.index += 1

        st.dataframe(
            top_df.style.format({
                "CTR%":                   "{:.2f}%",
                "TC%":                    "{:.2f}%",
                "Relative Lift":          "{:.1f}%",
                "Incremental Conversions":"{:.1f}",
                "Total Converted":        "{:,.0f}",
            }),
            use_container_width=True
        )

# ── Download ──────────────────────────────────────────────────────────────────
st.divider()

with open(output_file, "rb") as f:
    report_bytes = f.read()

st.download_button(
    label="📥 Download Full Report",
    data=report_bytes,
    file_name=os.path.basename(output_file),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary"
)
