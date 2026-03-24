import sys
import os
import re
import unicodedata
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


def find_first_matching_column(df, candidates):
    """Return first existing column from candidates (case-insensitive exact, then contains)."""
    if df is None or df.empty:
        return None
    normalized = {c.strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in normalized:
            return normalized[key]
    for candidate in candidates:
        key = candidate.strip().lower()
        for col in df.columns:
            low = col.strip().lower()
            if key in low:
                return col
    return None


def normalize_citizenship(series):
    """Keep uploaded citizenship labels as-is; only normalize missing to Unknown."""
    if series is None:
        return series
    return normalize_unknown(series)


def normalize_unknown(series):
    """Normalize blanks/missing values to 'Unknown' for reporting."""
    s = series.astype(str).str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "none": pd.NA, "N/A": pd.NA, "n/a": pd.NA})
    return s.fillna("Unknown")

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

# Invisible / odd space chars in exports (zero-width, BOM, NBSP) break naive "esl" == checks.
_ESL_STRIP_INVISIBLE = re.compile(r"[\u200b-\u200d\ufeff\u00a0]+")


def _normalized_esl_key(val):
    """Fold Unicode + strip invisibles so ESL, ＥＳＬ, E\u200bSL, etc. match."""
    if val is pd.NA or (isinstance(val, float) and pd.isna(val)):
        return None
    s = unicodedata.normalize("NFKC", str(val))
    s = _ESL_STRIP_INVISIBLE.sub("", s).strip()
    return s.casefold()


# Major / school text that implies ESL for Program Type (word token or full phrase).
_ESL_TOKEN_IN_TEXT = re.compile(r"(?<!\w)esl(?!\w)", re.IGNORECASE)
_ESL_LONG_PHRASE = re.compile(r"english\s+as\s+(?:a\s+)?second\s+language", re.IGNORECASE)


def _text_indicates_esl_program(val):
    """True if free text is ESL or describes English as a Second Language."""
    if val is None:
        return False
    try:
        if pd.isna(val):
            return False
    except (TypeError, ValueError):
        pass
    s = unicodedata.normalize("NFKC", str(val))
    s = _ESL_STRIP_INVISIBLE.sub("", s).strip()
    if not s:
        return False
    if _normalized_esl_key(s) == "esl":
        return True
    if _ESL_TOKEN_IN_TEXT.search(s):
        return True
    cf = re.sub(r"\s+", " ", s.casefold())
    if _ESL_LONG_PHRASE.search(cf):
        return True
    return False


def canonicalize_school_display_name(name):
    """School column: treat esl / ESL / Esl / esL (Unicode-folded) as one label 'ESL'."""
    if name is None:
        return name
    try:
        if pd.isna(name):
            return name
    except (TypeError, ValueError):
        pass
    s = str(name).strip()
    if not s:
        return s
    if _normalized_esl_key(s) == "esl":
        return "ESL"
    m = re.match(r"^\([^)]+\)\s*(.+)$", s)
    if m and _normalized_esl_key(m.group(1).strip()) == "esl":
        return "ESL"
    return s


def program_type_from_irregular_field(df):
    """
    Program type rules for reporting:
    - If Derived Academic Status is Never_Enrolled -> Unknown
    - Else if Irregular Program field is blank -> Regular
    - Else use the Irregular Program value (e.g., Option III); ESL matches case-insensitively as "ESL"
    - If Major or any Pseudo Sch cell indicates ESL / English as a Second Language, Program Type is ESL
      (overrides Irregular Program and Regular), except Never_Enrolled stays Unknown.
    """
    status_col = next(
        (c for c in df.columns if c.strip().lower() == "derived academic status"),
        None,
    )
    irregular_col = next(
        (c for c in df.columns if c.strip().lower() == "irregular program"),
        None,
    )

    # Start as Regular for enrolled students; convert missing labels to Unknown later.
    program = pd.Series(["Regular"] * len(df), index=df.index, dtype=object)

    if irregular_col:
        raw = df[irregular_col].astype(str).str.strip()
        raw = raw.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "none": pd.NA})
        esl_keys = raw.map(lambda x: _normalized_esl_key(x) if pd.notna(x) else pd.NA)
        esl_mask = raw.notna() & esl_keys.eq("esl")
        raw = raw.mask(esl_mask, "ESL")
        # If not blank, preserve category value from uploaded file (ESL normalized above).
        program = raw.fillna("Regular")

    never_mask = pd.Series(False, index=df.index)
    if status_col:
        status = df[status_col].astype(str).str.strip().str.lower()
        never_mask = status.eq("never_enrolled")
        program.loc[never_mask] = "Unknown"

    major_col = find_first_matching_column(df, ["Maj1 Name", "Major", "Major Name"])
    major_esl = (
        df[major_col].map(_text_indicates_esl_program)
        if major_col
        else pd.Series(False, index=df.index)
    )
    pseudo_cols = [c for c in df.columns if c.strip().lower().startswith("pseudo sch")]
    school_esl = pd.Series(False, index=df.index)
    for c in pseudo_cols:
        school_esl = school_esl | df[c].map(_text_indicates_esl_program)

    esl_from_major_or_school = major_esl | school_esl
    program.loc[esl_from_major_or_school & ~never_mask] = "ESL"

    return normalize_unknown(program)


# Embedded school lookup: script looks for this file in script dir, CSV dir, or Downloads
SCHOOL_LOOKUP_FILENAME = "COLA Toolkit, Spring 2026.csv"
# Enrollment reference for representation comparison (optional)
ENROLLMENT_REFERENCE_FILENAME = "All_International_Students_Enrolled.csv"

# Built‑in code → school mapping (used if no external lookup file is present).
DEFAULT_SCHOOL_CODE_LOOKUP = {
    "2": "Business Administration",
    "3": "Education",
    "4": "Engineering",
    "5": "Fine Arts",
    "6": "Graduate School",
    "7": "Law School",
    "8": "Pharmacy",
    "9": "Architecture",
    "B": "Graduate Business",
    "C": "Communication",
    "E": "Natural Sciences",
    "F": "Civic Leadership",
    "J": "Geosciences",
    "L": "Liberal Arts",
    "M": "Medical School",
    "N": "Nursing",
    "P": "Information",
    "S": "Social Work",
    "T": "Public Affairs",
    "U": "Undergraduate Studies",
}


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
    """Load school code -> official school name mapping from a lookup CSV, layered over defaults."""
    mapping = dict(DEFAULT_SCHOOL_CODE_LOOKUP)
    if not path:
        return mapping
    try:
        lut = pd.read_csv(path)
    except FileNotFoundError:
        return mapping

    lut.columns = lut.columns.str.strip()
    if not {"Code", "School"}.issubset(lut.columns):
        print(f"School lookup file {path} is missing 'Code' or 'School' columns. Using built‑in school code mapping.")
        return mapping

    lut = lut[
        (lut["Code"].astype(str).str.strip() != "")
        & (lut["School"].astype(str).str.strip() != "")
    ]
    csv_mapping = dict(
        zip(
            lut["Code"].astype(str).str.strip(),
            lut["School"].astype(str).str.strip(),
        )
    )
    mapping.update(csv_mapping)
    return mapping


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
            result.add(canonicalize_school_display_name(code_to_school.get(code, s)))
        else:
            result.add(canonicalize_school_display_name(s))
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

    # Be flexible about pseudo school column names (e.g. 'Pseudo Sch1', 'Pseudo Sch Maj1', etc.).
    pseudo_cols = [
        c for c in df.columns if c.strip().lower().startswith("pseudo sch")
    ]
    if not pseudo_cols:
        return pd.Series(dtype=object), pd.DataFrame(columns=["Category", "Count", "Proportion"])

    col1 = pseudo_cols[0]
    col2 = pseudo_cols[1] if len(pseudo_cols) > 1 else None
    student_school_counts = Counter()
    for idx, row in df.iterrows():
        schools = _parse_schools_from_cell(row[col1], code_to_school)
        if col2:
            schools |= _parse_schools_from_cell(row[col2], code_to_school)
        valid_schools = [s for s in schools if s and not _is_graduate_school(s)]
        for s in valid_schools:
            if s and not _is_graduate_school(s):
                student_school_counts[s] += 1
        if not valid_schools:
            student_school_counts["Unknown"] += 1

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
            enr["School"] = enr["School"].map(canonicalize_school_display_name)
            enr = enr[enr["School"].astype(str).str.strip() != ""]
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

    If a school appears twice with and without a leading code, e.g.
    "(2)Business Administration" and "Business Administration", they are
    merged into a single row using the plain-school label.
    """
    if total_participants <= 0 or total_enrollment <= 0:
        return pd.DataFrame(
            columns=[
                "School",
                "Enrollment Count",
                "Enrollment %",
                "Participant Count",
                "Participation %",
                "Representation Ratio",
            ]
        )

    def _normalize_label(label: str) -> str:
        s = str(label).strip()
        m = re.match(r"\(([^)]+)\)\s*(.+)", s)
        core = m.group(2).strip() if m else s
        return canonicalize_school_display_name(core)

    # Normalize and merge participation counts
    part_norm = {}
    for school, val in participation_counts.items():
        key = _normalize_label(school)
        part_norm[key] = part_norm.get(key, 0) + int(val)

    # Normalize and merge enrollment counts
    enr_norm = {}
    for school, val in enrollment_counts.items():
        key = _normalize_label(school)
        enr_norm[key] = enr_norm.get(key, 0) + int(val)

    all_schools = sorted(set(part_norm.keys()) | set(enr_norm.keys()))
    rows = []
    for school in all_schools:
        enc = int(enr_norm.get(school, 0))
        prc = int(part_norm.get(school, 0))
        # Skip rows where both enrollment and participation are zero
        if enc == 0 and prc == 0:
            continue
        enr_pct = (enc / total_enrollment) * 100 if total_enrollment > 0 else 0.0
        part_pct = (prc / total_participants) * 100 if total_participants > 0 else 0.0
        ratio = (part_pct / enr_pct) if enr_pct else float("nan")
        rows.append(
            {
                "School": school,
                "Enrollment Count": enc,
                "Enrollment %": enr_pct,
                "Participant Count": prc,
                "Participation %": part_pct,
                "Representation Ratio": ratio,
            }
        )
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
    """Remove invalid student rows (empty/missing names). Keep Never_Enrolled rows for Program Type = Unknown."""
    n_before = len(df)
    # Drop rows with no name (empty or blank)
    if "Name" in df.columns:
        df = df[df["Name"].astype(str).str.strip().str.lower() != "nan"]
        df = df[df["Name"].astype(str).str.strip() != ""]
    n_after = len(df)
    removed = n_before - n_after
    if removed > 0:
        print(f"Cleaning: removed {removed} row(s) with missing student identity. Analyzed {n_after} students.")
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
    school_counts, school_proportions_df = student_based_school_proportions(df, code_to_school)

    base = os.path.splitext(os.path.basename(event_csv_path))[0]
    out_dir = os.path.dirname(os.path.abspath(event_csv_path))
    out_docx = os.path.join(out_dir, f"{base}_report.docx")

    doc = Document()
    doc.add_heading('Event Participant Proportions Report', 0)
    doc.add_paragraph(f"Source: {event_csv_path}")
    doc.add_paragraph(
        f"All tables and charts in this report are based on N = {enrolled_n} enrolled student participant(s) "
        "from the Advisor Toolkit export. Rows with missing names are excluded. In Program Type only, "
        "students marked 'Never_Enrolled' are shown as Unknown."
    )
    # Build list of categories from the uploaded participant CSV only.
    categories = []
    major_col = find_first_matching_column(df, ["Maj1 Name", "Major", "Major Name"])
    if major_col:
        categories.append((normalize_unknown(df[major_col]), "Proportion of Majors"))

    # Explicit demographic charts requested: Gender and Citizenship
    gender_col = find_first_matching_column(df, ["Gender"])
    if gender_col:
        categories.append((normalize_unknown(df[gender_col]), "Proportion of Gender"))

    citizenship_col = find_first_matching_column(df, ["Citizenship"])
    if citizenship_col:
        categories.append((normalize_unknown(normalize_citizenship(df[citizenship_col])), "Proportion of Citizenship"))

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
    categories.append((normalize_unknown(program_series), "Proportion of Regular vs Irregular Programs"))

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
