"""Streamlit web interface for the gradebook tool."""

import json
from io import BytesIO

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
    st.subheader("2 · Config")
    config_tab, editor_tab = st.tabs(["Upload config.json", "Paste / edit JSON"])

    with config_tab:
        config_file = st.file_uploader(
            "config.json or config.yaml",
            type=["json", "yaml", "yml"],
            key="config_upload",
        )

    with editor_tab:
        default_config = json.dumps(
            {
                "course": "JM204 Media and Social Diversity",
                "semester": "2025-2",
                "professors": ["Dr. Smith"],
                "grade_scale": {
                    "A": 90, "B+": 85, "B": 80,
                    "C+": 75, "C": 70, "D+": 65, "D": 60,
                },
                "pass_threshold": 50,
                "categories": [
                    {
                        "name": "Assignments",
                        "weight": 30,
                        "raw_max": 32,
                        "assignments": [
                            "Attendance & Quiz",
                            "Homework #1 - Questioning Media Reality",
                        ],
                    },
                    {
                        "name": "Midterm",
                        "weight": 25,
                        "raw_max": 25,
                        "assignments": ["Midterm - Individual Media Analysis"],
                    },
                    {
                        "name": "Final Project",
                        "weight": 35,
                        "raw_max": 35,
                        "assignments": ["FINAL PROJECT - Individual Submission"],
                    },
                    {"name": "Participation", "weight": 10, "manual": True},
                ],
            },
            indent=2,
        )
        config_text = st.text_area(
            "Edit config JSON",
            value=default_config,
            height=340,
            key="config_text",
        )

st.divider()

# ── Options ───────────────────────────────────────────────────────────────────
with st.expander("Advanced options"):
    fuzzy_threshold = st.slider(
        "Fuzzy-match threshold",
        min_value=30,
        max_value=100,
        value=60,
        step=5,
        help=(
            "Minimum similarity score (0–100) to match an export's assignment name to a "
            "config category. Lower = more lenient. Unmatched assignments go to 'Unassigned'."
        ),
    )
    output_filename = st.text_input(
        "Output filename",
        value="gradebook.xlsx",
        help="Name for the downloaded .xlsx file.",
    )

# ── Generate ──────────────────────────────────────────────────────────────────
st.subheader("3 · Generate")
generate = st.button("⚙️  Build Gradebook", type="primary", disabled=not uploaded_files)

if generate:
    # Resolve config
    config: dict | None = None
    try:
        if config_file is not None:
            config = load_config(config_file)
        else:
            config = json.loads(config_text)
    except Exception as exc:
        st.error(f"Config parse error: {exc}")
        st.stop()

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
        file_name=output_filename,
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

**Config fields:**
| Field | Description |
|---|---|
| `course` | Course name shown in the gradebook banner |
| `semester` | Semester label |
| `professors` | List of instructor names (supports multiple) |
| `grade_scale` | Letter → minimum score mapping |
| `pass_threshold` | Minimum total score to pass |
| `categories[].name` | Category name |
| `categories[].weight` | Percentage weight (should sum to 100) |
| `categories[].raw_max` | Raw points total for this category (for scaling) |
| `categories[].assignments` | Assignment names to include (fuzzy matched) |
| `categories[].manual` | `true` for Participation — leaves column blank |

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
