"""Streamlit web interface for the gradebook tool."""

import json

import streamlit as st

try:
    from core import load_config, process
except Exception as _import_err:
    st.error(f"Failed to import core module: {_import_err}")
    st.stop()

st.set_page_config(
    page_title="Gradebook Builder",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }
    .stDownloadButton > button {
        background-color: #1F3864;
        color: white;
        font-weight: bold;
        font-size: 1.1rem;
        padding: 0.6rem 2rem;
        border-radius: 8px;
    }
    .warning-box {
        background: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 0.75rem 1rem;
        border-radius: 4px;
        margin: 0.5rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📊 Gradebook Builder")
st.caption("MS Teams export → weighted .xlsx gradebook with Excel formulas")

st.divider()

# ── Layout ────────────────────────────────────────────────────────────────────
left, right = st.columns([1, 1], gap="large")

with left:
    st.subheader("1 · Upload Assignment Exports")
    uploaded_files = st.file_uploader(
        "Drag & drop MS Teams .xlsx export files",
        type=["xlsx"],
        accept_multiple_files=True,
        help=(
            "One file per assignment. Header must be on row 2 (index 1). "
            "Required columns: Full Name, Assignments, Points."
        ),
    )
    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) uploaded")
        with st.expander("File list"):
            for f in uploaded_files:
                st.markdown(f"• `{f.name}`")

with right:
    st.subheader("2 · Course settings")

    course_name = st.text_input("Course name", value="JM204 Media and Social Diversity")
    col_sem, col_prof = st.columns(2)
    with col_sem:
        semester = st.text_input("Semester", value="2025-2")
    with col_prof:
        professors_raw = st.text_input(
            "Instructor(s)", value="Dr. Smith",
            help="Separate multiple names with a comma",
        )

    st.markdown("**Grade categories** — one per row: Name | Weight% | Max raw points | Assignment names (comma-separated)")
    st.caption("Leave 'Max raw points' blank and tick 'Manual entry' for Participation.")

    # Default category rows
    default_categories = [
        ("Assignments", 30, 32, "Attendance & Quiz, Homework #1 - Questioning Media Reality", False),
        ("Midterm", 25, 25, "Midterm - Individual Media Analysis", False),
        ("Final Project", 35, 35, "FINAL PROJECT - Individual Submission", False),
        ("Participation", 10, 0, "", True),
    ]

    categories = []
    for i, (dname, dweight, draw_max, dasgns, dmanual) in enumerate(default_categories):
        with st.expander(f"Category {i+1}: {dname}", expanded=True):
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                cat_name = st.text_input("Name", value=dname, key=f"cat_name_{i}")
            with c2:
                cat_weight = st.number_input("Weight %", value=dweight, min_value=0, max_value=100, key=f"cat_weight_{i}")
            with c3:
                cat_manual = st.checkbox("Manual entry", value=dmanual, key=f"cat_manual_{i}",
                                         help="Tick for Participation — leaves the column blank for you to fill in")
            if not cat_manual:
                cat_raw_max = st.number_input("Max raw points total", value=draw_max, min_value=1, key=f"cat_raw_{i}")
                cat_asgns_raw = st.text_area(
                    "Assignment names (one per line, fuzzy matched)",
                    value="\n".join(a.strip() for a in dasgns.split(",") if a.strip()),
                    height=80,
                    key=f"cat_asgns_{i}",
                    help="Paste the assignment names exactly as they appear in your MS Teams exports. Partial matches are fine.",
                )
                asgn_list = [a.strip() for a in cat_asgns_raw.splitlines() if a.strip()]
                categories.append({"name": cat_name, "weight": cat_weight,
                                    "raw_max": cat_raw_max, "assignments": asgn_list})
            else:
                categories.append({"name": cat_name, "weight": cat_weight, "manual": True})

    with st.expander("Grade scale"):
        gc1, gc2, gc3, gc4 = st.columns(4)
        grade_A   = gc1.number_input("A ≥",  value=90, key="gA")
        grade_Bp  = gc1.number_input("B+ ≥", value=85, key="gBp")
        grade_B   = gc2.number_input("B ≥",  value=80, key="gB")
        grade_Cp  = gc2.number_input("C+ ≥", value=75, key="gCp")
        grade_C   = gc3.number_input("C ≥",  value=70, key="gC")
        grade_Dp  = gc3.number_input("D+ ≥", value=65, key="gDp")
        grade_D   = gc4.number_input("D ≥",  value=60, key="gD")
        pass_threshold = gc4.number_input("Pass threshold ≥", value=50, key="gPass",
                                          help="Students below this total are marked Fail")

st.divider()

# ── Generate ──────────────────────────────────────────────────────────────────
st.subheader("3 · Generate")
generate = st.button("⚙️  Build Gradebook", type="primary", disabled=not uploaded_files)

if generate:
    config: dict = {
        "course": course_name,
        "semester": semester,
        "professors": [p.strip() for p in professors_raw.split(",") if p.strip()],
        "grade_scale": {
            "A": grade_A, "B+": grade_Bp, "B": grade_B,
            "C+": grade_Cp, "C": grade_C, "D+": grade_Dp, "D": grade_D,
        },
        "pass_threshold": pass_threshold,
        "categories": categories,
    }
    with st.spinner("Parsing exports and building gradebook …"):
        try:
            xlsx_bytes, unmatched = process(uploaded_files, config)
        except Exception as exc:
            st.error(f"Processing failed: {exc}")
            st.stop()

    if unmatched:
        st.markdown(
            "<div class='warning-box'>"
            "<strong>⚠️ Unmatched assignments</strong> — these did not match any config category "
            "and were placed in an <em>Unassigned</em> column:<br/>"
            + "".join(f"• {n}<br/>" for n in unmatched)
            + "</div>",
            unsafe_allow_html=True,
        )

    st.success("Gradebook ready!")
    st.download_button(
        label="⬇️  Download Gradebook",
        data=xlsx_bytes,
        file_name="gradebook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ── Help ──────────────────────────────────────────────────────────────────────
with st.expander("How to use"):
    st.markdown(
        """
**MS Teams export format expected:**
- Row 1 (index 0): may be a meta row (ignored)
- Row 2 (index 1): column headers — must include **Full Name**, **Assignments**, **Points**
- Row 3+: student data

**Output sheet 1 "Master Gradebook":**
- Row 1: Course banner
- Row 2: Colour-coded category headers
- Row 3: Column headers with max points
- Rows 4+: One student per row, alphabetical
- Red cells = missing or zero scores
- All totals, weighted scores, and letter grades are **Excel formulas** — adjust raw scores and the sheet recalculates automatically
- Participation column is left blank for manual entry
- Freeze panes at B4 so names stay visible when scrolling

**Output sheet 2 "Grading Key":** Category weights, grade scale, pass conditions.
        """
    )
