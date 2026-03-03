import os
import tempfile

import streamlit as st

from Generate_Proportion import generate_report


st.set_page_config(page_title="Event Participant Proportions Report", layout="centered")

st.title("Event Participant Proportions Report")
st.write(
    "Upload your event participants CSV (and optionally an enrollment reference CSV). "
    "The app will generate a Word report with tables and charts, similar to the desktop script."
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
)

st.markdown(
    """
### Enrollment reference CSV (for comparison to overall enrollment)

If you provide this, the report will include enrollment vs participation by school.

- **Option 1 – Summary format:**
  - `School`
  - One of: `Enrollment`, `Count`, or `Students`
- **Option 2 – Student-level format** (same structure as event CSV):
  - `Pseudo Sch1` (and optionally `Pseudo Sch2`)

If you leave this blank, the script will try to use
`All_International_Students_Enrolled.csv` from the event CSV folder, the script folder,
or your Downloads folder.
"""
)

enrollment_file = st.file_uploader(
    "Upload enrollment reference CSV (optional)",
    type=["csv"],
    accept_multiple_files=False,
    help="Leave blank to let the script auto-detect All_International_Students_Enrolled.csv.",
)

# Single Generate button
generate_clicked = st.button("Generate report", key="generate_report_button")

if generate_clicked:
    if event_file is None:
        st.warning("Please upload an event participants CSV first.")
        st.stop()
    with st.spinner("Generating report..."):
        # Save uploaded files to temporary directory
        with tempfile.TemporaryDirectory() as tmpdir:
            event_path = os.path.join(tmpdir, event_file.name)
            with open(event_path, "wb") as f:
                f.write(event_file.read())

            enrollment_path = None
            if enrollment_file is not None:
                enrollment_path = os.path.join(tmpdir, enrollment_file.name)
                with open(enrollment_path, "wb") as f:
                    f.write(enrollment_file.read())

            try:
                report_path = generate_report(event_path, enrollment_path)
            except Exception as e:
                st.error(f"Error while generating report: {e}")
            else:
                # Read report bytes and offer download
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
