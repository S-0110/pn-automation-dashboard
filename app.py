import streamlit as st
import subprocess

st.title("PN Automation Dashboard")

uploaded_file = st.file_uploader(
    "Upload Raw File",
    type=["xlsx"]
)

if uploaded_file:

    # Save uploaded file
    with open("1_to_10.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("Run Automation"):

        result = subprocess.run(
            ["python", "pn.py"],
            capture_output=True,
            text=True
        )

        st.text(result.stdout)

        output_file = "Campaign_Analysis_Output.xlsx"

        with open(output_file, "rb") as f:

            st.download_button(
                label="Download Report",
                data=f,
                file_name="Campaign_Analysis_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
