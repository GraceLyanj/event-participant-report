import os
import tempfile

import streamlit as st

from Generate_Proportion import generate_report
from Extract_EIDs import extract_eids_from_directory, write_eids_to_csv


st.set_page_config(page_title="Event Participant Proportions Report", layout="centered")

st.title("Event Participant Proportions Report")

tab_report, tab_extract_eids = st.tabs(["Generate report", "Extract EIDs from multiple CSVs"])

with tab_report:
    st.write(
        "Upload your event participants CSV. The app will generate a Word report with tables and charts, "
        "using embedded enrollment data (All_International_Students_Enrolled.csv) for comparison."
    )

    st.markdown(
        """
### Event participants CSV

- **Required for cleaning and proportions:**
  - `Name`
  - `Derived Academic Status`
  - `Pseudo Sch1`
- **Recommended (for more tables/charts):**
  - `Pseudo Sch2` (if students can have a second school)
  - `Maj1 Name`
  - `Gender`
  - `Citizenship`
  - `Irregular Program`
"""
    )

    event_file = st.file_uploader(
        "Upload event participants CSV",
        type=["csv"],
        accept_multiple_files=False,
        help="This should be the CSV you currently pass to Generate_Proportion.py.",
        key="event_file",
    )

    generate_clicked = st.button("Generate report", key="generate_report_button")

    if generate_clicked:
        if event_file is None:
            st.warning("Please upload an event participants CSV first.")
            st.stop()
        with st.spinner("Generating report..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                event_path = os.path.join(tmpdir, event_file.name)
                with open(event_path, "wb") as f:
                    f.write(event_file.read())

                try:
                    report_path = generate_report(event_path, None)
                except Exception as e:
                    st.error(f"Error while generating report: {e}")
                else:
                    with open(report_path, "rb") as f:
                        report_bytes = f.read()

                    report_name = os.path.basename(report_path)
                    st.success("Report generated successfully.")
                    st.download_button(
                        label="Download report (.docx)",
                        data=report_bytes,
                        file_name=report_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

with tab_extract_eids:
    st.write(
        "Upload one or more CSV files from a folder. All EIDs from EID-like columns "
        "will be collected and merged into a single list. After extraction, you can copy a newline‑separated list "
        "of unique EIDs from the box below (with a copy icon)."
    )
    st.markdown("Uses the same logic as **Extract_EIDs.py** (one folder = one set of CSVs here).")

    csv_uploads = st.file_uploader(
        "Upload CSV file(s)",
        type=["csv"],
        accept_multiple_files=True,
        help="Select all CSV files.",
        key="extract_eids_uploads",
    )

    extract_clicked = st.button("Extract EIDs", key="extract_eids_button")

    if extract_clicked:
        if not csv_uploads:
            st.warning("Please upload at least one CSV file.")
        else:
            with st.spinner("Extracting EIDs from all CSVs..."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    for up in csv_uploads:
                        path = os.path.join(tmpdir, up.name)
                        with open(path, "wb") as f:
                            f.write(up.read())
                    try:
                        eids = extract_eids_from_directory(tmpdir)
                    except Exception as e:
                        st.error(f"Error while extracting EIDs: {e}")
                    else:
                        if not eids:
                            st.info("No EIDs found in any of the uploaded CSVs.")
                        else:
                            # Show EIDs as a newline‑separated, directly copyable block
                            eids_list = sorted(eids)
                            eids_text = "\n".join(eids_list)
                            st.success(f"Found **{len(eids)}** unique EID(s). Copy them from the box below.")
                            st.code(eids_text, language="text")
