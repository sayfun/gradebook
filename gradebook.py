"""Streamlit web interface for the gradebook tool."""

import streamlit as st

try:
    from core import parse_teams_export, process, match_assignment_to_category
except Exception as _import_err:
    st.error(f"Failed to import core module: {_import_err}")
    st.stop()

st.set_page_config(
    page_title="Gradebook Builder",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }
    .stDownloadButton > button {
        background-color: #1F3864; color: white;
        font-weight: bold; font-size: 1.1rem;
        padding: 0.6rem 2rem; border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Gradebook Builder")
st.caption("MS Teams exports → weighted .xlsx gradebook with Excel formulas")
st.divider()

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.subheader("1 · Upload assignment exports")
uploaded_files = st.file_uploader(
    "Drag & drop all MS Teams .xlsx export files (one file per assignment)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload your MS Teams export files above to get started.")
    st.stop()

st.success(f"{len(uploaded_files)} file(s) uploaded")

# Parse assignment names from every file (cached on the file set)
@st.cache_data(show_spinner=False)
def detect_assignments(file_contents: list[tuple[str, bytes]]) -> list[str]:
    """Return sorted list of unique assignment names detected across all exports."""
    from io import BytesIO
    names = []
    for fname, data in file_contents:
        try:
            buf = BytesIO(data)
            buf.name = fname
            asgn_name, _ = parse_teams_export(buf)
            if asgn_name and asgn_name not in names:
                names.append(asgn_name)
        except Exception:
            pass
    return sorted(names)

file_contents = [(f.name, f.read()) for f in uploaded_files]
# Reset file pointers for later use
for f in uploaded_files:
    f.seek(0)

detected = detect_assignments(file_contents)

if not detected:
    st.error("Could not detect any assignment names from the uploaded files. Check that the files are valid MS Teams exports.")
    st.stop()

st.divider()

# ── Step 2: Course settings + categories ──────────────────────────────────────
st.subheader("2 · Course settings")

col_a, col_b, col_c = st.columns(3)
course_name    = col_a.text_input("Course name", value="JM204 Media and Social Diversity")
semester       = col_b.text_input("Semester", value="2025-2")
professors_raw = col_c.text_input("Instructor(s)", value="Dr. Smith",
                                   help="Separate multiple names with a comma")

st.markdown("**Grade categories**")
st.caption("Add as many categories as you need. Participation (or any manual category) gets a blank column you fill in yourself.")

# ── Dynamic category list stored in session state ─────────────────────────────
if "categories" not in st.session_state:
    st.session_state.categories = [
        {"name": "Assignments",   "weight": 30, "raw_max": 32, "manual": False},
        {"name": "Midterm",       "weight": 25, "raw_max": 25, "manual": False},
        {"name": "Final Project", "weight": 35, "raw_max": 35, "manual": False},
        {"name": "Participation", "weight": 10, "raw_max": 10, "manual": True},
    ]

def add_category():
    st.session_state.categories.append(
        {"name": f"Category {len(st.session_state.categories)+1}",
         "weight": 0, "raw_max": 10, "manual": False}
    )

def remove_category(idx):
    st.session_state.categories.pop(idx)

for i, cat in enumerate(st.session_state.categories):
    with st.expander(f"**{cat['name']}**  —  {cat['weight']}%", expanded=False):
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
        st.session_state.categories[i]["name"] = c1.text_input(
            "Category name", value=cat["name"], key=f"cname_{i}")
        st.session_state.categories[i]["weight"] = c2.number_input(
            "Weight %", value=cat["weight"], min_value=0, max_value=100, key=f"cw_{i}")
        st.session_state.categories[i]["manual"] = c3.checkbox(
            "Manual entry", value=cat["manual"], key=f"cman_{i}",
            help="Tick this for Participation — leaves the column blank for you to fill in after export")
        if not st.session_state.categories[i]["manual"]:
            st.session_state.categories[i]["raw_max"] = c4.number_input(
                "Max raw pts", value=cat.get("raw_max", 10), min_value=1, key=f"crm_{i}")
        if st.button("Remove", key=f"rem_{i}"):
            remove_category(i)
            st.rerun()

st.button("＋ Add category", on_click=add_category)

total_weight = sum(c["weight"] for c in st.session_state.categories)
if total_weight != 100:
    st.warning(f"Weights sum to {total_weight}% — should be 100%.")

with st.expander("Grade scale"):
    gc1, gc2, gc3, gc4 = st.columns(4)
    grade_A  = gc1.number_input("A ≥",  value=90, key="gA")
    grade_Bp = gc1.number_input("B+ ≥", value=85, key="gBp")
    grade_B  = gc2.number_input("B ≥",  value=80, key="gB")
    grade_Cp = gc2.number_input("C+ ≥", value=75, key="gCp")
    grade_C  = gc3.number_input("C ≥",  value=70, key="gC")
    grade_Dp = gc3.number_input("D+ ≥", value=65, key="gDp")
    grade_D  = gc4.number_input("D ≥",  value=60, key="gD")
    pass_thr = gc4.number_input("Pass threshold ≥", value=50, key="gPass")

st.divider()

# ── Step 3: Assign detected assignments to categories ─────────────────────────
st.subheader("3 · Assign assignments to categories")
st.caption(
    "Every assignment detected from your uploaded files is listed below. "
    "Pick which grade category it belongs to — or skip it if it shouldn't appear in the gradebook."
)

non_manual_cats = [c["name"] for c in st.session_state.categories if not c.get("manual")]
category_options = non_manual_cats + ["— skip this assignment —"]

# Build a quick config stub for fuzzy-match suggestions
_stub_config = {
    "categories": [
        {**c, "assignments": []}
        for c in st.session_state.categories
        if not c.get("manual")
    ]
}

assignment_map: dict[str, str] = {}  # assignment name → chosen category name

for asgn in detected:
    # Suggest best fuzzy match as default
    suggested = match_assignment_to_category(asgn, _stub_config, threshold=50)
    default_idx = (
        category_options.index(suggested)
        if suggested and suggested in category_options
        else len(category_options) - 1  # "skip"
    )
    chosen = st.selectbox(
        f"**{asgn}**",
        options=category_options,
        index=default_idx,
        key=f"asgn_{asgn}",
    )
    if chosen != "— skip this assignment —":
        assignment_map[asgn] = chosen

st.divider()

# ── Step 4: Generate ──────────────────────────────────────────────────────────
st.subheader("4 · Generate")

assigned_count = len(assignment_map)
skipped_count  = len(detected) - assigned_count
st.caption(f"{assigned_count} assignment(s) assigned · {skipped_count} skipped")

if st.button("⚙️  Build Gradebook", type="primary"):
    # Build category list with assignments attached
    cat_asgn_map: dict[str, list[str]] = {c["name"]: [] for c in st.session_state.categories}
    for asgn, cat in assignment_map.items():
        cat_asgn_map[cat].append(asgn)

    categories_cfg = []
    for cat in st.session_state.categories:
        if cat.get("manual"):
            categories_cfg.append({"name": cat["name"], "weight": cat["weight"], "manual": True})
        else:
            categories_cfg.append({
                "name": cat["name"],
                "weight": cat["weight"],
                "raw_max": cat.get("raw_max", 10),
                "assignments": cat_asgn_map.get(cat["name"], []),
            })

    config = {
        "course": course_name,
        "semester": semester,
        "professors": [p.strip() for p in professors_raw.split(",") if p.strip()],
        "grade_scale": {
            "A": grade_A, "B+": grade_Bp, "B": grade_B,
            "C+": grade_Cp, "C": grade_C, "D+": grade_Dp, "D": grade_D,
        },
        "pass_threshold": pass_thr,
        "categories": categories_cfg,
    }

    # Re-supply file contents (file pointers were consumed by detect_assignments)
    from io import BytesIO
    sources = [BytesIO(data) for _, data in file_contents]
    for src, (fname, _) in zip(sources, file_contents):
        src.name = fname

    with st.spinner("Building gradebook …"):
        try:
            xlsx_bytes, unmatched = process(sources, config)
        except Exception as exc:
            st.error(f"Processing failed: {exc}")
            st.stop()

    if unmatched:
        st.warning(
            "These assignments didn't match any category and were placed in an "
            f"**Unassigned** column: {', '.join(unmatched)}"
        )

    st.success("Gradebook ready!")
    st.download_button(
        label="⬇️  Download Gradebook",
        data=xlsx_bytes,
        file_name="gradebook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
