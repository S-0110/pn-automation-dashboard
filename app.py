import streamlit as st
import subprocess
import os

st.title("PN Automation Dashboard")

uploaded_file = st.file_uploader(
    "Upload Raw File",
    type=["xlsx"]
)

if uploaded_file:

    with open("1_to_10.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("Run Automation"):

        with st.spinner("Running Automation..."):

            result = subprocess.run(
                ["python", "pn.py"],
                capture_output=True,
                text=True
            )

        st.write("Return Code:", result.returncode)

        if result.stdout:
            st.code(result.stdout)

        if result.stderr:
            st.error(result.stderr)

        output_file = "Campaign_Analysis_Output.xlsx"

        if os.path.exists(output_file):

            st.success("Report Generated Successfully")

            with open(output_file, "rb") as f:

                st.download_button(
                    label="Download Report",
                    data=f,
                    file_name="Campaign_Analysis_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        else:
            st.error("Output file was not generated.")
