import sys
import os
import re
from collections import Counter
# Check required packages first and give a clear fix if missing
def _check_imports():
    missing = []
    try:
        import pandas
    except ImportError:
        missing.append("pandas")
    try:
        import matplotlib
    except ImportError:
        missing.append("matplotlib")
    try:
        import docx
    except ImportError:
        missing.append("python-docx")
    if missing:
        pip_names = {"pandas": "pandas", "matplotlib": "matplotlib", "python-docx": "python-docx", "docx": "python-docx"}
        to_install = [pip_names.get(m, m) for m in missing]
        to_install = list(dict.fromkeys(to_install))  # dedupe (docx -> python-docx)
        exe = sys.executable
        print("Missing required package(s):", ", ".join(missing))
        print("Using Python:", exe)
        print("Run this in the same terminal you use to run the script:")
        print(f'  "{exe}" -m pip install {" ".join(to_install)}')
        sys.exit(1)

_check_imports()

import pandas as pd
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def get_proportions_df(series):
    """Return (counts, proportions, dataframe) for the series."""
    counts = series.value_counts(dropna=False)
    proportions = counts / counts.sum()
    df = pd.DataFrame({'Category': counts.index.astype(str), 'Count': counts.values, 'Proportion': proportions.values})
    return counts, proportions, df

def add_table_to_doc(doc, title, df, style='Table Grid'):
    """Add a formatted Word table with a title."""
    doc.add_paragraph(title, style='Heading 2')
    nrows, ncols = len(df) + 1, len(df.columns)
    table = doc.add_table(rows=nrows, cols=ncols)
    table.style = style
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header row
    for j, col in enumerate(df.columns):
        cell = table.rows[0].cells[j]
        cell.text = col
        _shade_cell(cell, 'D9E2F3')
    # Data rows
    for i, row in enumerate(df.itertuples(index=False)):
        for j, val in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            if df.columns[j] == 'Proportion':
                cell.text = f"{val:.1%}"
            else:
                cell.text = str(val)
    doc.add_paragraph()

def _shade_cell(cell, fill_hex):
    """Apply light shading to a table cell."""
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill_hex)
    cell._tc.get_or_add_tcPr().append(shd)

def program_type_from_irregular_field(df):
    """If 'Irregular Program' has any value, student is Irregular (IP); otherwise Regular."""
    if "Irregular Program" not in df.columns:
        return pd.Series(["Regular"] * len(df), index=df.index)
    return df["Irregular Program"].astype(str).str.strip().replace("nan", "").replace("", pd.NA).notna().map({True: "Irregular", False: "Regular"})


# Embedded school lookup: script looks for this file in script dir, CSV dir, or Downloads
SCHOOL_LOOKUP_FILENAME = "COLA Toolkit, Spring 2026.csv"
# Enrollment reference for representation comparison (optional)
ENROLLMENT_REFERENCE_FILENAME = "All_International_Students_Enrolled.csv"


def resolve_school_lookup_path(csv_dir, script_dir=None):
    """Return path to school lookup file if it exists in script dir, CSV dir, or user Downloads."""
    candidates = []
    if script_dir:
        candidates.append(os.path.join(script_dir, SCHOOL_LOOKUP_FILENAME))
    candidates.append(os.path.join(csv_dir, SCHOOL_LOOKUP_FILENAME))
    downloads = os.path.join(os.path.expanduser("~"), "Downloads", SCHOOL_LOOKUP_FILENAME)
    candidates.append(downloads)
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def resolve_enrollment_path(csv_dir, script_dir=None, explicit_path=None):
    """Return path to enrollment reference CSV if provided or found (csv_dir, script_dir, Downloads)."""
    if explicit_path and os.path.isfile(explicit_path):
        return explicit_path
    candidates = []
    if script_dir:
        candidates.append(os.path.join(script_dir, ENROLLMENT_REFERENCE_FILENAME))
    candidates.append(os.path.join(csv_dir, ENROLLMENT_REFERENCE_FILENAME))
    downloads = os.path.join(os.path.expanduser("~"), "Downloads", ENROLLMENT_REFERENCE_FILENAME)
    candidates.append(downloads)
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def load_school_lookup(path):
    """Load school code -> official school name mapping from a lookup CSV."""
    try:
        lut = pd.read_csv(path)
    except FileNotFoundError:
        return {}

    lut.columns = lut.columns.str.strip()
    if not {"Code", "School"}.issubset(lut.columns):
        print(f"School lookup file {path} is missing 'Code' or 'School' columns. Using pseudo school names as-is.")
        return {}

    lut = lut[
        (lut["Code"].astype(str).str.strip() != "")
        & (lut["School"].astype(str).str.strip() != "")
    ]
    return dict(
        zip(
            lut["Code"].astype(str).str.strip(),
            lut["School"].astype(str).str.strip(),
        )
    )


def translate_pseudo_school(series, code_to_school):
    """Translate '(Code)School' pseudo values to official school names using lookup."""
    if series is None or not code_to_school:
        return series

    def _translate(val):
        if pd.isna(val):
            return val
        s = str(val)
        # Extract code inside parentheses, e.g. "(E)Natural Sciences" -> "E"
        m = re.search(r"\(([^)]+)\)", s)
        if not m:
            return s
        code = m.group(1).strip()
        return code_to_school.get(code, s)

    return series.apply(_translate)


def _is_graduate_school(school_name):
    """True if this school is Graduate School (not exclusive; excluded from school breakdown)."""
    return str(school_name).strip().lower() == "graduate school"


def _parse_schools_from_cell(val, code_to_school):
    """Parse one cell (may contain 'A/ B'). Return set of translated school names."""
    if pd.isna(val) or str(val).strip() == "":
        return set()
    parts = [p.strip() for p in str(val).split("/") if p.strip()]
    result = set()
    for s in parts:
        m = re.search(r"\(([^)]+)\)", s)
        if m:
            code = m.group(1).strip()
            result.add(code_to_school.get(code, s))
        else:
            result.add(s)
    return result


def student_based_school_proportions(df, code_to_school):
    """
    College/School proportions by student count: each student counts once in the denominator.
    A student in multiple schools (e.g. Graduate + COLA) is counted in each of those schools;
    proportions are (students in school X) / total_students, so they can sum to more than 100%.
    Uses Pseudo Sch1 and Pseudo Sch2; cells with 'A/ B' are split into multiple schools.
    Returns (counts_series, proportions_df) for table/chart.
    """
    total_students = len(df)
    if total_students == 0:
        return pd.Series(dtype=object), pd.DataFrame(columns=["Category", "Count", "Proportion"])

    col1 = "Pseudo Sch1"
    col2 = "Pseudo Sch2" if "Pseudo Sch2" in df.columns else None
    student_school_counts = Counter()
    for idx, row in df.iterrows():
        schools = _parse_schools_from_cell(row[col1], code_to_school)
        if col2:
            schools |= _parse_schools_from_cell(row[col2], code_to_school)
        for s in schools:
            if s and not _is_graduate_school(s):
                student_school_counts[s] += 1

    if not student_school_counts:
        return pd.Series(dtype=object), pd.DataFrame(columns=["Category", "Count", "Proportion"])

    counts = pd.Series(student_school_counts)
    proportions = counts / total_students
    proportions_df = pd.DataFrame({
        "Category": counts.index.astype(str),
        "Count": counts.values,
        "Proportion": proportions.values,
    })
    return counts, proportions_df


def load_enrollment_by_school(path, code_to_school):
    """
    Load enrollment counts by school from a CSV.
    Supports: (1) Student-level CSV with Pseudo Sch1 / Pseudo Sch2 (same as event data);
              (2) Summary CSV with 'School' and 'Enrollment' or 'Count'.
    Returns (series: school -> count, total_enrollment).
    """
    try:
        enr = pd.read_csv(path)
    except Exception:
        return pd.Series(dtype=object), 0
    enr.columns = enr.columns.str.strip()
    # Summary format: School + Enrollment, Count, or Students
    if "School" in enr.columns:
        count_col = next(
            (c for c in ("Enrollment", "Count", "Students") if c in enr.columns),
            None,
        )
        if count_col:
            enr = enr.dropna(subset=["School"])
            enr["School"] = enr["School"].astype(str).str.strip()
            enr = enr[enr["School"] != ""]
            total_enrollment = int(enr[count_col].sum())
            enr = enr[~enr["School"].str.lower().str.strip().eq("graduate school")]
            counts = enr.groupby("School", as_index=True)[count_col].sum()
            return counts, total_enrollment
    # Student-level format (same as event CSV)
    if "Pseudo Sch1" in enr.columns:
        counts_series, _ = student_based_school_proportions(enr, code_to_school)
        total = len(enr)
        return counts_series, total
    return pd.Series(dtype=object), 0


def build_representation_comparison(participation_counts, total_participants, enrollment_counts, total_enrollment):
    """
    Build comparison table: Enrollment %, Participation %, Representation Ratio.
    Ratio = (Participation %) / (Enrollment %). ≈1 proportional, >1 overrepresented, <1 underrepresented.
    """
    if total_participants <= 0 or total_enrollment <= 0:
        return pd.DataFrame(columns=["School", "Enrollment Count", "Enrollment %", "Participant Count", "Participation %", "Representation Ratio"])
    all_schools = sorted(set(participation_counts.index) | set(enrollment_counts.index))
    rows = []
    for school in all_schools:
        enc = int(enrollment_counts.get(school, 0))
        prc = int(participation_counts.get(school, 0))
        enr_pct = (enc / total_enrollment) * 100
        part_pct = (prc / total_participants) * 100
        ratio = (part_pct / enr_pct) if enr_pct else float("nan")
        rows.append({
            "School": school,
            "Enrollment Count": enc,
            "Enrollment %": enr_pct,
            "Participant Count": prc,
            "Participation %": part_pct,
            "Representation Ratio": ratio,
        })
    return pd.DataFrame(rows)


def comparison_table_to_doc(doc, df):
    """Add representation comparison table to document."""
    if df.empty:
        return
    doc.add_paragraph("Enrollment % = (School Enrollment / Total Enrollment) × 100. Participation % = (School Participants / Total Participants) × 100. Ratio = Participation % / Enrollment % (≈1 proportional, >1 overrepresented, <1 underrepresented).", style="Normal")
    nrows, ncols = len(df) + 1, len(df.columns)
    table = doc.add_table(rows=nrows, cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, col in enumerate(df.columns):
        cell = table.rows[0].cells[j]
        cell.text = col
        _shade_cell(cell, "D9E2F3")
    for i, row in enumerate(df.itertuples(index=False)):
        for j, val in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            if "Ratio" in df.columns[j]:
                cell.text = f"{val:.2f}" if pd.notna(val) else "—"
            elif "%" in df.columns[j]:
                cell.text = f"{val:.1f}%"
            else:
                cell.text = str(val)
    doc.add_paragraph()


def side_by_side_bar_chart_bytes(comparison_df, title="Enrollment % vs Participation % by School"):
    """Side-by-side bar chart: Enrollment % and Participation % per school."""
    if comparison_df.empty or "School" not in comparison_df.columns:
        buf = io.BytesIO()
        plt.figure(figsize=(8, 4))
        plt.text(0.5, 0.5, "No data", ha="center", va="center")
        plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
        plt.close()
        buf.seek(0)
        return buf
    df = comparison_df.sort_values("Enrollment %", ascending=False).head(16)
    x = range(len(df))
    w = 0.35
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar([i - w / 2 for i in x], df["Enrollment %"], width=w, label="Enrollment %", color="steelblue")
    ax.bar([i + w / 2 for i in x], df["Participation %"], width=w, label="Participation %", color="coral")
    ax.set_xticks(x)
    ax.set_xticklabels(df["School"].astype(str).str[:30] + df["School"].astype(str).str.len().gt(30).map({True: "…", False: ""}), rotation=45, ha="right")
    ax.set_ylabel("Percentage")
    ax.set_title(title)
    ax.legend()
    ax.axhline(y=0, color="gray", linewidth=0.5)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
    plt.close()
    buf.seek(0)
    return buf


def representation_ratio_chart_bytes(comparison_df, title="Representation Ratio & Over/Under by School"):
    """Bar chart of representation ratio (1 = proportional). Gray ≈ proportional, red under, blue/green over."""
    if comparison_df.empty or "Representation Ratio" not in comparison_df.columns:
        buf = io.BytesIO()
        plt.figure(figsize=(8, 4))
        plt.text(0.5, 0.5, "No data", ha="center", va="center")
        plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
        plt.close()
        buf.seek(0)
        return buf
    df = comparison_df.dropna(subset=["Representation Ratio"]).sort_values("Representation Ratio", ascending=True).tail(16)
    if df.empty:
        buf = io.BytesIO()
        plt.figure(figsize=(8, 4))
        plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
        plt.close()
        buf.seek(0)
        return buf
    colors = ["gray" if 0.98 <= r <= 1.02 else "tomato" if r < 1 else "forestgreen" for r in df["Representation Ratio"]]
    fig, ax = plt.subplots(figsize=(10, 5))
    y_pos = range(len(df))
    ax.barh(y_pos, df["Representation Ratio"], color=colors)
    ax.axvline(x=1.0, color="black", linestyle="--", linewidth=1, label="Proportional (1.0)")
    ax.set_yticks(y_pos)
    ax.set_yticklabels(df["School"].astype(str).str[:35] + df["School"].astype(str).str.len().gt(35).map({True: "…", False: ""}))
    ax.set_xlabel("Representation Ratio (gray ≈ proportional, green over, red under)")
    ax.set_title(title)
    ax.legend()
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
    plt.close()
    buf.seek(0)
    return buf


def _prepare_pie_counts(counts, max_slices=12):
    """Group small slices into 'Other' if there are too many."""
    if len(counts) <= max_slices:
        return counts
    top = counts.head(max_slices - 1)
    other_count = counts.iloc[max_slices - 1:].sum()
    other_label = f"Other ({len(counts) - max_slices + 1} categories)"
    return pd.concat([top, pd.Series({other_label: other_count})])

def pie_chart_to_bytes(counts, title, max_slices=12):
    """Draw a pie chart with legend (no labels on slices) to avoid overlapping text."""
    counts = _prepare_pie_counts(counts, max_slices)
    labels = [str(x)[:40] + ('...' if len(str(x)) > 40 else '') for x in counts.index]
    fig, ax = plt.subplots(figsize=(7, 5))
    wedges, texts, autotexts = ax.pie(
        counts.values,
        labels=None,
        autopct='%1.1f%%',
        startangle=90,
        pctdistance=0.6,
        explode=[0.02] * len(counts),
    )
    ax.set_title(title, fontsize=12)
    ax.legend(wedges, labels, title='Category', loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=120, bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf

def clean_unknown_students(df):
    """Remove unknown/invalid student rows: empty rows and Never_Enrolled (unknown roster) entries."""
    n_before = len(df)
    # Drop rows with no name (empty or blank)
    if "Name" in df.columns:
        df = df[df["Name"].astype(str).str.strip().str.lower() != "nan"]
        df = df[df["Name"].astype(str).str.strip() != ""]
    # Drop "Never_Enrolled" / unknown roster students
    if "Derived Academic Status" in df.columns:
        df = df[df["Derived Academic Status"].astype(str).str.strip().str.lower() != "never_enrolled"]
    n_after = len(df)
    removed = n_before - n_after
    if removed > 0:
        print(f"Cleaning: removed {removed} unknown or non-enrolled student row(s). Analyzed {n_after} students.")
    return df.reset_index(drop=True)

def _parse_never_enrolled_eids(raw_text):
    """
    Parse Advisor Toolkit messages like
    'dk33895 does not appear to have ever enrolled.'
    and return a sorted list of unique EIDs.
    """
    if not raw_text:
        return []
    eids = set()
    for line in str(raw_text).splitlines():
        line = line.strip()
        if not line:
            continue
        # Grab the first token before whitespace or '('
        m = re.match(r"([A-Za-z0-9]+)", line)
        if m:
            eids.add(m.group(1))
    return sorted(eids)


def generate_report(event_csv_path, enrollment_reference_path=None, never_enrolled_notes=None):
    """
    Generate the Word report for a given event participants CSV.

    Parameters
    ----------
    event_csv_path : str
        Path to the event participants CSV.
    enrollment_reference_path : str or None, optional
        Optional path to the enrollment reference CSV. If None, the script
        will look for All_International_Students_Enrolled.csv in the event
        CSV directory, script directory, or user Downloads.
    never_enrolled_notes : str or None, optional
        Optional raw text from Advisor Toolkit listing EIDs that
        "do not appear to have ever enrolled". These EIDs are described
        in the report as outside the enrolled-student sample and counted
        only as part of the irregular program environment; they are not
        added to the event participant dataset.

    Returns
    -------
    str
        Path to the generated .docx report.
    """
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv else os.getcwd()

    try:
        df = pd.read_csv(event_csv_path)
    except FileNotFoundError:
        print(f"File not found: {event_csv_path}")
        raise

    df.columns = df.columns.str.strip()

    # Clean before any analysis
    df = clean_unknown_students(df)
    if len(df) == 0:
        raise ValueError("No students remaining after cleaning; cannot generate report.")
    enrolled_n = len(df)

    # Parse "never enrolled" EIDs from optional notes
    never_enrolled_eids = _parse_never_enrolled_eids(never_enrolled_notes)

    # School proportions: by student count (graduate school can co-occur with other schools)
    csv_dir = os.path.dirname(os.path.abspath(event_csv_path))
    school_lookup_path = resolve_school_lookup_path(csv_dir, script_dir)
    code_to_school = load_school_lookup(school_lookup_path) if school_lookup_path else {}
    if "Pseudo Sch1" in df.columns:
        school_counts, school_proportions_df = student_based_school_proportions(df, code_to_school)
    else:
        school_counts, school_proportions_df = pd.Series(dtype=object), pd.DataFrame(columns=["Category", "Count", "Proportion"])

    base = os.path.splitext(os.path.basename(event_csv_path))[0]
    out_dir = os.path.dirname(os.path.abspath(event_csv_path))
    out_docx = os.path.join(out_dir, f"{base}_report.docx")

    doc = Document()
    doc.add_heading('Event Participant Proportions Report', 0)
    doc.add_paragraph(f"Source: {event_csv_path}")
    doc.add_paragraph(
        f"All tables and charts in this report are based on N = {enrolled_n} enrolled student participant(s) "
        "from the Advisor Toolkit export. Students with 'Never_Enrolled' status or missing names in the event "
        "file are excluded from all analyses."
    )
    if never_enrolled_eids:
        doc.add_paragraph(
            "Advisor Toolkit also reported the following EID(s) as not appearing to have ever enrolled. "
            "They are counted as part of the irregular program environment in the Regular vs Irregular "
            "Program breakdown, but they are not included in the other tables or charts below."
        )
        doc.add_paragraph(", ".join(never_enrolled_eids))

    # Build list of categories only for columns that actually exist
    categories = []
    column_title_pairs = [
        ("Maj1 Name", "Proportion of Majors"),
        ("Gender", "Proportion of Gender"),
        ("Citizenship", "Proportion of Citizenship"),
    ]
    for col, title in column_title_pairs:
        if col in df.columns:
            categories.append((df[col], title))

    # Irregular = something in 'Irregular Program' field (student from IP)
    df["Program Type"] = program_type_from_irregular_field(df)
    # For program type proportions, also count each "never enrolled" EID as Irregular in the
    # environment. They are only added to this Regular vs Irregular breakdown, not to the
    # other demographic tables/charts.
    if never_enrolled_eids:
        program_series = pd.concat(
            [
                df["Program Type"],
                pd.Series(
                    ["Irregular"] * len(never_enrolled_eids),
                    name="Program Type",
                ),
            ],
            ignore_index=True,
        )
    else:
        program_series = df["Program Type"]
    categories.append((program_series, "Proportion of Regular vs Irregular Programs"))

    # Add formatted tables (one per category), with explicit sample sizes
    doc.add_heading('Summary tables', level=1)
    if not school_proportions_df.empty:
        add_table_to_doc(
            doc,
            f"Proportion of College/School (by student count, N = {enrolled_n})",
            school_proportions_df,
        )
    for series, title in categories:
        # Sample size for this category: non-missing, non-blank entries
        valid = series.dropna().astype(str).str.strip()
        valid_n = (valid != "").sum()
        titled = f"{title} (N = {valid_n})"
        counts, proportions, tbl = get_proportions_df(series)
        add_table_to_doc(doc, titled, tbl)

    # Add pie chart for each category (labels in legend to avoid overlapping), with sample sizes
    doc.add_heading('Charts', level=1)
    if not school_counts.empty:
        chart_title = f"Proportion of College/School (by student count, N = {enrolled_n})"
        doc.add_heading(chart_title, level=2)
        doc.add_picture(
            pie_chart_to_bytes(school_counts, chart_title),
            width=Inches(5.5),
        )
    for series, title in categories:
        counts = series.value_counts(dropna=False)
        # Match heading/sample size used in tables
        valid = series.dropna().astype(str).str.strip()
        valid_n = (valid != "").sum()
        chart_title = f"{title} (N = {valid_n})"
        doc.add_heading(chart_title, level=2)
        doc.add_picture(
            pie_chart_to_bytes(counts, chart_title),
            width=Inches(5.5),
        )

    # Comparison to international enrollment (if reference file provided or found)
    enrollment_path = resolve_enrollment_path(csv_dir, script_dir, enrollment_reference_path)
    if enrollment_path and not school_counts.empty:
        enrollment_counts, total_enrollment = load_enrollment_by_school(enrollment_path, code_to_school)
        if total_enrollment > 0 and not enrollment_counts.empty:
            total_participants = enrolled_n
            comparison_df = build_representation_comparison(
                school_counts, total_participants, enrollment_counts, total_enrollment
            )
            doc.add_heading("Comparison to International Enrollment (by School)", level=1)
            doc.add_paragraph(f"Enrollment reference: {enrollment_path}")
            comparison_table_to_doc(doc, comparison_df)
            doc.add_heading("Enrollment % vs Participation %", level=2)
            doc.add_picture(side_by_side_bar_chart_bytes(comparison_df), width=Inches(5.5))
            doc.add_heading("Representation Ratio & Over/Under by School", level=2)
            doc.add_picture(representation_ratio_chart_bytes(comparison_df, title="Representation Ratio & Over/Under by School"), width=Inches(5.5))
        else:
            print("Enrollment file found but no enrollment counts by school; skipping comparison section.")
    elif not school_counts.empty:
        print("No enrollment reference file found; comparison section omitted. Place All_International_Students_Enrolled.csv in the CSV dir or pass it as second argument.")

    try:
        doc.save(out_docx)
        print(f"Report saved to: {out_docx}")
        return out_docx
    except PermissionError:
        from datetime import datetime
        alt_name = f"{base}_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        alt_path = os.path.join(out_dir, alt_name)
        doc.save(alt_path)
        print(f"Original file is open or locked; report saved to: {alt_path}")
        return alt_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python Generate_Proportion.py <event_participants.csv> [enrollment_reference.csv]")
        print("  If enrollment_reference.csv is omitted, looks for All_International_Students_Enrolled.csv in the same dir, script dir, or Downloads.")
        sys.exit(1)

    filename = sys.argv[1]
    enrollment_explicit = sys.argv[2] if len(sys.argv) > 2 else None
    generate_report(filename, enrollment_explicit)

if __name__ == "__main__":
    main()
