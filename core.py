"""Core gradebook processing logic shared by CLI and web app."""

from __future__ import annotations

import json
import re
from io import BytesIO
from pathlib import Path
from typing import IO, Union

import pandas as pd
import yaml
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from thefuzz import fuzz

# ── colour palette ──────────────────────────────────────────────────────────
CATEGORY_COLOURS = [
    "4472C4",  # blue
    "ED7D31",  # orange
    "A9D18E",  # green
    "FFC000",  # gold
    "9DC3E6",  # light blue
    "FF0000",  # red (fallback)
]
RED_FILL = PatternFill("solid", fgColor="FFCCCC")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")


# ── config loading ───────────────────────────────────────────────────────────

def load_config(path: Union[str, Path, IO]) -> dict:
    """Load JSON or YAML config from a file path or file-like object."""
    if hasattr(path, "read"):
        raw = path.read()
        if isinstance(raw, bytes):
            raw = raw.decode()
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return yaml.safe_load(raw)

    path = Path(path)
    text = path.read_text(encoding="utf-8")
    if path.suffix in (".yaml", ".yml"):
        return yaml.safe_load(text)
    return json.loads(text)


# ── Teams export parser ──────────────────────────────────────────────────────

def parse_teams_export(
    source: Union[str, Path, IO],
    *,
    header_row: int = 1,
) -> tuple[str, dict[str, float]]:
    """
    Parse one MS Teams .xlsx export.

    Returns (assignment_name, {student_name: points}).
    Header is at row index `header_row` (0-based), data starts next row.
    """
    df = pd.read_excel(source, header=header_row, engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Locate the required columns (case-insensitive, partial match)
    col_map: dict[str, str] = {}
    for col in df.columns:
        lc = col.lower()
        if "full name" in lc or lc == "name":
            col_map["name"] = col
        elif "assignment" in lc and "name" not in lc and "assignment" not in col_map:
            col_map["assignment"] = col
        elif lc == "points" or lc.startswith("point"):
            col_map["points"] = col
        elif "max" in lc and "point" in lc:
            col_map["max_points"] = col

    required = {"name", "points"}
    missing = required - col_map.keys()
    if missing:
        raise ValueError(
            f"Could not find columns {missing} in export. "
            f"Available columns: {list(df.columns)}"
        )

    # Derive assignment name: prefer the "Assignments" column value from first
    # non-empty row; fall back to filename stem.
    assignment_name: str = ""
    if "assignment" in col_map:
        vals = df[col_map["assignment"]].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            assignment_name = vals.iloc[0]

    if not assignment_name and hasattr(source, "name"):
        assignment_name = Path(source.name).stem
    elif not assignment_name and isinstance(source, (str, Path)):
        assignment_name = Path(source).stem

    scores: dict[str, float] = {}
    for _, row in df.iterrows():
        student = str(row[col_map["name"]]).strip() if pd.notna(row[col_map["name"]]) else ""
        if not student or student.lower() in ("nan", ""):
            continue
        pts = row[col_map["points"]]
        if pd.isna(pts):
            scores[student] = float("nan")
        else:
            try:
                scores[student] = float(pts)
            except (ValueError, TypeError):
                scores[student] = float("nan")

    return assignment_name, scores


# ── fuzzy assignment → category matching ─────────────────────────────────────

def _strip(s: str) -> str:
    """Normalise for fuzzy matching: lowercase, remove emoji & punctuation."""
    s = s.lower().strip()
    # remove emoji
    s = re.sub(r"[^\x00-\x7F]+", " ", s)
    # collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s


def match_assignment_to_category(
    assignment_name: str,
    config: dict,
    threshold: int = 60,
) -> str | None:
    """
    Return the category name for an assignment or None if unmatched.

    Uses token-set ratio so 'Homework #1 - Questioning Media Reality' and
    'Homework 1 Questioning Media Reality' both match.
    """
    a_norm = _strip(assignment_name)
    best_score = 0
    best_category: str | None = None

    for cat in config["categories"]:
        if cat.get("manual"):
            continue
        for candidate in cat.get("assignments", []):
            c_norm = _strip(candidate)
            score = fuzz.token_set_ratio(a_norm, c_norm)
            if score > best_score:
                best_score = score
                best_category = cat["name"]

    if best_score >= threshold:
        return best_category
    return None


# ── merge all exports ────────────────────────────────────────────────────────

class MergeResult:
    def __init__(
        self,
        df: pd.DataFrame,
        assignment_categories: dict[str, str],
        unmatched: list[str],
    ):
        self.df = df                            # student × assignment
        self.assignment_categories = assignment_categories  # assignment → category
        self.unmatched = unmatched              # assignment names with no category


def merge_exports(
    sources: list[Union[str, Path, IO]],
    config: dict,
) -> MergeResult:
    """
    Parse all export files and merge into one wide DataFrame.

    Index = student name, columns = assignment names.
    """
    all_scores: dict[str, dict[str, float]] = {}  # assignment → {student → score}
    assignment_categories: dict[str, str] = {}
    unmatched: list[str] = []

    for src in sources:
        try:
            asgn_name, scores = parse_teams_export(src)
        except Exception as exc:
            name = getattr(src, "name", str(src))
            raise ValueError(f"Failed to parse '{name}': {exc}") from exc

        if asgn_name in all_scores:
            # Merge (later file wins where both have a value)
            existing = all_scores[asgn_name]
            for student, score in scores.items():
                if student not in existing or (pd.isna(existing[student]) and not pd.isna(score)):
                    existing[student] = score
        else:
            all_scores[asgn_name] = scores

        if asgn_name not in assignment_categories:
            cat = match_assignment_to_category(asgn_name, config)
            if cat is None:
                if asgn_name not in unmatched:
                    unmatched.append(asgn_name)
                assignment_categories[asgn_name] = "Unassigned"
            else:
                assignment_categories[asgn_name] = cat

    # Union of all students
    all_students: set[str] = set()
    for scores in all_scores.values():
        all_students.update(scores.keys())
    students = sorted(all_students)

    # Build DataFrame
    records = []
    for student in students:
        row: dict[str, object] = {"Full Name": student}
        for asgn, scores in all_scores.items():
            row[asgn] = scores.get(student, float("nan"))
        records.append(row)

    df = pd.DataFrame(records).set_index("Full Name")
    return MergeResult(df, assignment_categories, unmatched)


# ── Excel formula builders ───────────────────────────────────────────────────

def _grade_formula(total_col: str, grade_scale: dict, row: int) -> str:
    """Build nested IF formula for letter grade from grade_scale thresholds."""
    # Sort descending by threshold
    sorted_grades = sorted(grade_scale.items(), key=lambda x: -x[1])
    ref = f"{total_col}{row}"
    formula = '"F"'
    for letter, threshold in reversed(sorted_grades):
        formula = f'IF({ref}>={threshold},"{letter}",{formula})'
    return f"={formula}"


def _category_formula(
    student_row: int,
    assignment_cols: list[str],
    raw_max: float,
    weight: float,
) -> str:
    """=SUM(C4:E4)/raw_max*weight  — all cell refs, no hardcoded sums."""
    if len(assignment_cols) == 1:
        cell_range = f"{assignment_cols[0]}{student_row}"
    else:
        cell_range = f"{assignment_cols[0]}{student_row}:{assignment_cols[-1]}{student_row}"
    return f"=IFERROR(SUM({cell_range})/{raw_max}*{weight},\"\")"


def _total_formula(student_row: int, cat_score_cols: list[str], participation_col: str) -> str:
    """Sum all category weighted score columns including participation."""
    parts = [f"{c}{student_row}" for c in cat_score_cols + [participation_col]]
    return "=IFERROR(" + "+".join(parts) + ',"")'


def _avg_formula(col: str, first_data_row: int, last_data_row: int) -> str:
    return f"=IFERROR(AVERAGE({col}{first_data_row}:{col}{last_data_row}),\"\")"


# ── Excel writer ─────────────────────────────────────────────────────────────

def build_excel(result: MergeResult, config: dict) -> bytes:
    """
    Produce a formatted .xlsx gradebook as bytes.
    All grade calculations are Excel formulas.
    """
    df = result.df
    assignment_categories = result.assignment_categories
    grade_scale = config["grade_scale"]
    pass_threshold = config.get("pass_threshold", 50)

    # ── lay out columns ──────────────────────────────────────────────────────
    # Order: non-manual categories in config order, then Participation, then
    # any Unassigned assignments.

    config_cats = [c for c in config["categories"] if not c.get("manual")]
    manual_cats = [c for c in config["categories"] if c.get("manual")]

    # For each non-manual category: ordered list of assignment column names
    cat_assignment_cols: dict[str, list[str]] = {c["name"]: [] for c in config_cats}
    unassigned_cols: list[str] = []

    for asgn in df.columns:
        cat = assignment_categories.get(asgn, "Unassigned")
        if cat == "Unassigned":
            unassigned_cols.append(asgn)
        elif cat in cat_assignment_cols:
            cat_assignment_cols[cat].append(asgn)

    # Build column sequence
    # Each entry: (display_name, kind, category, excel_col_letter [filled later])
    columns: list[dict] = [{"name": "Full Name", "kind": "name", "category": None}]

    cat_colour_map: dict[str, str] = {}
    for i, cat_cfg in enumerate(config_cats):
        cat_name = cat_cfg["name"]
        cat_colour_map[cat_name] = CATEGORY_COLOURS[i % len(CATEGORY_COLOURS)]
        for asgn in cat_assignment_cols[cat_name]:
            columns.append({"name": asgn, "kind": "assignment", "category": cat_name})
        # Category weighted score column
        columns.append({
            "name": f"{cat_name} /{cat_cfg['weight']}",
            "kind": "cat_score",
            "category": cat_name,
            "cat_cfg": cat_cfg,
        })

    # Manual participation columns
    for cat_cfg in manual_cats:
        cat_colour_map[cat_cfg["name"]] = CATEGORY_COLOURS[
            (len(config_cats)) % len(CATEGORY_COLOURS)
        ]
        columns.append({
            "name": f"{cat_cfg['name']} /{cat_cfg['weight']}",
            "kind": "participation",
            "category": cat_cfg["name"],
            "cat_cfg": cat_cfg,
        })

    # Unassigned
    for asgn in unassigned_cols:
        columns.append({"name": asgn, "kind": "unassigned", "category": "Unassigned"})

    # Total and grade
    columns.append({"name": "Total /100", "kind": "total", "category": None})
    columns.append({"name": "Letter Grade", "kind": "grade", "category": None})
    columns.append({"name": "Pass/Fail", "kind": "passfail", "category": None})

    # Assign excel column letters
    for i, col in enumerate(columns):
        col["letter"] = get_column_letter(i + 1)

    # ── helper lookups ───────────────────────────────────────────────────────
    def col_letter(kind=None, category=None, name=None) -> str:
        for c in columns:
            if name and c["name"] == name:
                return c["letter"]
            if kind and c["kind"] == kind and (category is None or c["category"] == category):
                return c["letter"]
        raise KeyError(f"Column not found: kind={kind} cat={category} name={name}")

    total_col = col_letter(kind="total")
    grade_col = col_letter(kind="grade")
    passfail_col = col_letter(kind="passfail")

    # ── row layout ───────────────────────────────────────────────────────────
    ROW_TITLE = 1
    ROW_CAT_HEADER = 2
    ROW_COL_HEADER = 3
    ROW_DATA_START = 4

    students = list(df.index)
    n_students = len(students)
    ROW_DATA_END = ROW_DATA_START + n_students - 1
    ROW_AVERAGE = ROW_DATA_END + 2
    ROW_DIST_START = ROW_AVERAGE + 3

    # ── workbook ─────────────────────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "Master Gradebook"
    ws2 = wb.create_sheet("Grading Key")

    n_cols = len(columns)
    last_col_letter = get_column_letter(n_cols)

    # ── Row 1: title banner ──────────────────────────────────────────────────
    title_parts = [config.get("course", ""), config.get("semester", "")]
    professors = config.get("professors", [])
    if professors:
        title_parts.append("Instructors: " + " & ".join(professors))
    title_text = "  |  ".join(p for p in title_parts if p)

    ws.merge_cells(f"A{ROW_TITLE}:{last_col_letter}{ROW_TITLE}")
    cell = ws.cell(ROW_TITLE, 1, title_text)
    cell.font = Font(bold=True, size=14, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F3864")
    cell.alignment = CENTER
    ws.row_dimensions[ROW_TITLE].height = 24

    # ── Row 2: category group headers ────────────────────────────────────────
    # Build contiguous spans per category by scanning columns left-to-right.
    # A category may appear in non-contiguous positions (e.g. None for "Full Name",
    # "Total", "Grade"), so we find contiguous runs and merge each run separately.
    def _contiguous_runs(col_list: list[dict]) -> list[tuple[dict, dict]]:
        """Return (first_col, last_col) tuples for each contiguous run."""
        runs: list[tuple[dict, dict]] = []
        run_start = col_list[0]
        prev_idx = columns.index(col_list[0])
        for c in col_list[1:]:
            idx = columns.index(c)
            if idx == prev_idx + 1:
                prev_idx = idx
            else:
                runs.append((run_start, columns[prev_idx]))
                run_start = c
                prev_idx = idx
        runs.append((run_start, columns[prev_idx]))
        return runs

    # Group columns by category preserving order
    cat_columns: dict[str | None, list[dict]] = {}
    for col in columns:
        cat_columns.setdefault(col["category"], []).append(col)

    for cat, cat_cols in cat_columns.items():
        colour = cat_colour_map.get(cat, "888888") if cat else "2C3E50"
        label = cat or ""
        for run_start, run_end in _contiguous_runs(cat_cols):
            start_letter = run_start["letter"]
            end_letter = run_end["letter"]
            col_idx = columns.index(run_start) + 1
            if start_letter != end_letter:
                ws.merge_cells(f"{start_letter}{ROW_CAT_HEADER}:{end_letter}{ROW_CAT_HEADER}")
            cell = ws.cell(ROW_CAT_HEADER, col_idx, label)
            cell.fill = PatternFill("solid", fgColor=colour)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = CENTER
            cell.border = THIN_BORDER

    ws.row_dimensions[ROW_CAT_HEADER].height = 20

    # ── Row 3: column headers ────────────────────────────────────────────────
    for i, col in enumerate(columns):
        cell = ws.cell(ROW_COL_HEADER, i + 1, col["name"])
        cat = col["category"]
        colour = cat_colour_map.get(cat, "4F4F4F") if cat else "4F4F4F"
        cell.fill = PatternFill("solid", fgColor=colour)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    ws.row_dimensions[ROW_COL_HEADER].height = 36

    # ── Data rows ────────────────────────────────────────────────────────────
    # Pre-compute which cols hold raw assignment scores per category for formula
    cat_raw_cols: dict[str, list[str]] = {}
    cat_score_col: dict[str, str] = {}
    participation_col_letter: str | None = None

    for col in columns:
        if col["kind"] == "assignment":
            cat_raw_cols.setdefault(col["category"], []).append(col["letter"])
        elif col["kind"] == "cat_score":
            cat_score_col[col["category"]] = col["letter"]
        elif col["kind"] == "participation":
            participation_col_letter = col["letter"]

    all_cat_score_cols = list(cat_score_col.values())

    for r_offset, student in enumerate(students):
        row_num = ROW_DATA_START + r_offset
        row_data = df.loc[student]

        # Name
        cell = ws.cell(row_num, 1, student)
        cell.alignment = LEFT
        cell.border = THIN_BORDER

        for col in columns:
            if col["kind"] == "name":
                continue
            col_idx = columns.index(col) + 1
            cell = ws.cell(row_num, col_idx)
            cell.border = THIN_BORDER
            cell.alignment = CENTER

            if col["kind"] == "assignment":
                val = row_data.get(col["name"])
                if pd.isna(val):
                    cell.fill = RED_FILL
                else:
                    cell.value = val
                    if val == 0:
                        cell.fill = RED_FILL

            elif col["kind"] == "cat_score":
                cat_cfg = col["cat_cfg"]
                asgn_letters = cat_raw_cols.get(col["category"], [])
                if asgn_letters:
                    cell.value = _category_formula(
                        row_num,
                        asgn_letters,
                        cat_cfg["raw_max"],
                        cat_cfg["weight"],
                    )
                    cell.number_format = "0.00"
                cat = col["category"]
                colour = cat_colour_map.get(cat)
                if colour:
                    cell.fill = PatternFill("solid", fgColor=_lighten(colour))

            elif col["kind"] == "participation":
                # Blank — instructor fills manually
                cat = col["category"]
                colour = cat_colour_map.get(cat)
                if colour:
                    cell.fill = PatternFill("solid", fgColor=_lighten(colour))

            elif col["kind"] == "unassigned":
                val = row_data.get(col["name"])
                if pd.isna(val):
                    cell.fill = RED_FILL
                else:
                    cell.value = val

            elif col["kind"] == "total":
                all_score_cols = all_cat_score_cols[:]
                part_col = participation_col_letter or ""
                cell.value = _total_formula(row_num, all_score_cols, part_col)
                cell.number_format = "0.00"
                cell.font = Font(bold=True)

            elif col["kind"] == "grade":
                cell.value = _grade_formula(total_col, grade_scale, row_num)
                cell.font = Font(bold=True)

            elif col["kind"] == "passfail":
                cell.value = (
                    f'=IF(ISNUMBER({total_col}{row_num}),'
                    f'IF({total_col}{row_num}>={pass_threshold},"Pass","Fail"),"")'
                )

    # ── Average row ───────────────────────────────────────────────────────────
    ws.cell(ROW_AVERAGE, 1, "Class Average").font = Font(bold=True)
    ws.cell(ROW_AVERAGE, 1).border = THIN_BORDER

    for col in columns[1:]:
        col_idx = columns.index(col) + 1
        cell = ws.cell(ROW_AVERAGE, col_idx)
        cell.border = THIN_BORDER
        if col["kind"] in ("assignment", "cat_score", "total"):
            cell.value = _avg_formula(col["letter"], ROW_DATA_START, ROW_DATA_END)
            cell.number_format = "0.00"
            cell.font = Font(bold=True)

    # ── Grade distribution ───────────────────────────────────────────────────
    dist_grades = sorted(grade_scale.keys(), key=lambda g: -grade_scale[g]) + ["F"]
    ws.cell(ROW_DIST_START, 1, "Grade Distribution").font = Font(bold=True, size=11)
    ws.cell(ROW_DIST_START, 2, "Count").font = Font(bold=True)
    ws.cell(ROW_DIST_START, 3, "Percent").font = Font(bold=True)
    for i, letter in enumerate(dist_grades):
        r = ROW_DIST_START + 1 + i
        ws.cell(r, 1, letter)
        grade_col_letter_for_count = grade_col
        count_formula = (
            f'=COUNTIF({grade_col_letter_for_count}{ROW_DATA_START}:'
            f'{grade_col_letter_for_count}{ROW_DATA_END},"{letter}")'
        )
        ws.cell(r, 2, count_formula)
        ws.cell(r, 3, f"=IFERROR(B{r}/{n_students},\"\")").number_format = "0.0%"

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 28
    for col in columns[1:]:
        if col["kind"] in ("cat_score", "total", "grade", "passfail", "participation"):
            ws.column_dimensions[col["letter"]].width = 14
        else:
            ws.column_dimensions[col["letter"]].width = 22

    # ── Freeze panes at B4 ───────────────────────────────────────────────────
    ws.freeze_panes = "B4"

    # ── Sheet 2: Grading Key ─────────────────────────────────────────────────
    _write_grading_key(ws2, config)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _lighten(hex_colour: str, factor: float = 0.6) -> str:
    """Return a lighter tint of a hex colour for background fills."""
    r, g, b = int(hex_colour[0:2], 16), int(hex_colour[2:4], 16), int(hex_colour[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"{r:02X}{g:02X}{b:02X}"


def _write_grading_key(ws, config: dict):
    ws.title = "Grading Key"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 30

    r = 1
    ws.cell(r, 1, "Grading Key").font = Font(bold=True, size=13)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor="1F3864")
    ws.cell(r, 1).font = Font(bold=True, size=13, color="FFFFFF")
    ws.merge_cells(f"A{r}:C{r}")
    r += 2

    ws.cell(r, 1, "Category").font = Font(bold=True)
    ws.cell(r, 2, "Weight").font = Font(bold=True)
    ws.cell(r, 3, "Raw Max").font = Font(bold=True)
    r += 1
    for cat in config["categories"]:
        ws.cell(r, 1, cat["name"])
        ws.cell(r, 2, f"{cat['weight']}%")
        ws.cell(r, 3, cat.get("raw_max", "manual"))
        r += 1
    r += 1

    ws.cell(r, 1, "Grade Scale").font = Font(bold=True)
    ws.cell(r, 2, "Min Score").font = Font(bold=True)
    r += 1
    scale = config["grade_scale"]
    for letter, threshold in sorted(scale.items(), key=lambda x: -x[1]):
        ws.cell(r, 1, letter)
        ws.cell(r, 2, f">= {threshold}")
        r += 1
    ws.cell(r, 1, "F")
    ws.cell(r, 2, f"< {min(scale.values())}")
    r += 2

    ws.cell(r, 1, "Pass Threshold").font = Font(bold=True)
    ws.cell(r, 2, f">= {config.get('pass_threshold', 50)}")
    r += 2

    if config.get("professors"):
        ws.cell(r, 1, "Instructors").font = Font(bold=True)
        ws.cell(r, 2, ", ".join(config["professors"]))


# ── public entry point ────────────────────────────────────────────────────────

def process(
    sources: list[Union[str, Path, IO]],
    config: Union[dict, str, Path, IO],
) -> tuple[bytes, list[str]]:
    """
    High-level entry point used by both CLI and web app.

    Returns (xlsx_bytes, unmatched_assignment_names).
    """
    if not isinstance(config, dict):
        config = load_config(config)

    result = merge_exports(sources, config)
    xlsx_bytes = build_excel(result, config)
    return xlsx_bytes, result.unmatched
