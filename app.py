import os
import tempfile
import base64

import streamlit as st
import streamlit.components.v1 as components
from docx import Document

from Generate_Proportion import generate_report
from Extract_EIDs import extract_eids_from_directory, write_eids_to_csv

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None


st.set_page_config(page_title="Event Participant Report", layout="centered")

# Always show the copy icon on code blocks (not just on hover)
st.markdown(
    """
    <style>
    [data-testid="stCodeBlock"] button {
        opacity: 1 !important;
        visibility: visible !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Event Participant Proportions Report")

tab_step1, tab_extract_eids, tab_report = st.tabs(
    [
        "Step 1 - Pull from Eventbrite",
        "Step 2 - Extract EIDs from multiple CSVs",
        "Step 3 - Generate report",
    ]
)

with tab_step1:
    st.markdown(
        """
### Step 1 - Pull data from Eventbrite

1. Go to the event you want to analyze in Eventbrite.
2. In the left sidebar, open **Reporting → Custom Question Responses**.
3. Under **Attendee Status**, select **Attended**.
4. In **Configure Columns**, deselect all columns, then select only **Custom Questions Responses**.
5. Click **Update Report**.
6. Export the report as a CSV file.
7. Repeat these steps for each event if you are analyzing an event series.
"""
    )

with tab_report:
    st.write(
        "Upload the CSV file you downloaded from Advisor Toolkit. The app will generate a Word report with tables and charts, "
        "using enrollment data (All_International_Students_Enrolled.csv) for comparison."
    )

    st.markdown(
        """
### Event participants CSV

- **Required for cleaning and proportions:**
  - `Name`
  - `Derived Academic Status`
  - `Pseudo Sch1`
- **Recommended (for more tables/charts):**
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

                    # Try to convert the .docx report to PDF for preview and download
                    pdf_bytes = None
                    pdf_name = os.path.splitext(report_name)[0] + ".pdf"
                    if docx2pdf_convert is not None:
                        try:
                            pdf_path = os.path.splitext(report_path)[0] + ".pdf"
                            docx2pdf_convert(report_path, pdf_path)
                            with open(pdf_path, "rb") as pf:
                                pdf_bytes = pf.read()
                        except Exception:
                            pdf_bytes = None

                    st.success("Report generated successfully.")
                    col_docx, col_pdf = st.columns(2)
                    with col_docx:
                        st.download_button(
                            label="Download report (.docx)",
                            data=report_bytes,
                            file_name=report_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    with col_pdf:
                        if pdf_bytes is not None:
                            st.download_button(
                                label="Download report (.pdf)",
                                data=pdf_bytes,
                                file_name=pdf_name,
                                mime="application/pdf",
                            )
                        else:
                            st.info("PDF download is currently unavailable; only the Word version can be downloaded.")

                    # Prefer a PDF preview if available, otherwise fall back to a simple text preview
                    if pdf_bytes is not None:
                        try:
                            b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
                            pdf_html = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="700" type="application/pdf"></iframe>'
                            st.markdown("#### Report preview (PDF)")
                            components.html(pdf_html, height=700)
                        except Exception:
                            pass
                    else:
                        try:
                            doc = Document(report_path)
                            preview_lines = []
                            for para in doc.paragraphs:
                                text = para.text.strip()
                                if text:
                                    preview_lines.append(text)
                                if len(preview_lines) >= 40:
                                    break
                            if preview_lines:
                                st.markdown("#### Report preview")
                                st.text_area(
                                    "First part of the report (preview)",
                                    value="\n".join(preview_lines),
                                    height=300,
                                )
                        except Exception:
                            # If preview fails for any reason, don't block downloads
                            pass

with tab_extract_eids:
    st.markdown(
        """
### Step 2 - Use Advisor Toolkit with Eventbrite CSVs

1. Open the **Advisor Toolkit** at [`https://utdirect.utexas.edu/link1/adtoolkit.WBX`](https://utdirect.utexas.edu/link1/adtoolkit.WBX).
2. In the left navigation bar, go to **Reporting toolkit → Latest data for EIDs (no semesters required)**.
3. Upload the CSV file(s) you exported from Eventbrite in Step 1 **here** to generate a clean list of EIDs using the tool below.
"""
    )

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
                            # Show EIDs as a newline‑separated, directly copyable code block
                            eids_list = sorted(eids)
                            eids_text = "\n".join(eids_list)
                            st.success(f"Found **{len(eids)}** unique EID(s). Use the copy button on the code block to copy them.")
                            st.code(
                                eids_text,
                                language="text",
                            )

    st.markdown(
        """
4. After you copy the EID list, use it in Advisor Toolkit to pull the latest data.
5. In Advisor Toolkit, include at least these fields in your report: **Major**, **Pseudo School(s)**, **Gender**, **Citizenship** (US citizen, PR, or international), and **Irregular Program** (e.g., Option III).
6. Generate the report and download it as a CSV.
"""
    )
