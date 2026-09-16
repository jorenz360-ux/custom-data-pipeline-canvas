"""
Best-effort parser for an existing gradebook Excel file, so the instructor
doesn't have to manually recreate the category/slot structure by hand for
every course. It reads the same structural conventions we found in the
IM101A Section D workbook:

  Row 1: big title merge
  Row 2: section header merges (ATTENDANCE / Activities / Quiz /
         Performance Task / Final Exam / Course Grade, ...)
  Row 3-4: per-column sub-headers -- either a raw-input label (a date for
         attendance, "Q2"/"PT1"/blank for a single raw score cell) or a
         computed-output label ("CA TOTAL", "%", "...GRADE")
  Row 5: "TOTAL POINTS" reference row
  Row 6+: one row per student

This is written generically -- keyword + structural heuristics, not
hardcoded column letters -- so it can be pointed at a different course's
sheet with a different layout (different section widths, missing section
titles, more or fewer raw columns per category) and still produce a
reasonable first pass. It NEVER writes to the database -- it only proposes
"detected slots" for the instructor to review on the Import page.
"""
import re
from datetime import datetime

import openpyxl
from openpyxl.utils import get_column_letter

CATEGORY_PATTERNS = [
    ("attendance", [r"attend"]),
    ("pt", [r"performance\s*task", r"\bpt\b"]),
    ("quiz", [r"quiz"]),
    ("exam", [r"exam"]),
    ("activity", [r"activit", r"\bca\b", r"class\s*standing"]),
]
OUTPUT_ONLY_PATTERNS = [
    r"^total$", r"^%$", r"grade", r"^status$", r"pass", r"fail", r"course\s*grade",
]
STUDENTS_RE = re.compile(r"\bstudents?\b", re.I)


def _text(ws, row, col):
    v = ws.cell(row=row, column=col).value
    return v.strip() if isinstance(v, str) else None


def _is_formula(value):
    if isinstance(value, str) and value.startswith("="):
        return True
    return type(value).__name__ == "ArrayFormula"


def _matches_any(text, patterns):
    if not text:
        return False
    t = text.lower()
    return any(re.search(p, t) for p in patterns)


def _match_category(text):
    if not text or _matches_any(text, OUTPUT_ONLY_PATTERNS):
        return None
    for cat, patterns in CATEGORY_PATTERNS:
        if _matches_any(text, patterns):
            return cat
    return None


def _guess_term(sheet_name):
    name = sheet_name.lower()
    if "mid" in name:
        return "midterm"
    if "final" in name:
        return "finals"
    return None


def _find_header_anchor(ws, max_scan_rows=10, max_scan_cols=10):
    for r in range(1, max_scan_rows + 1):
        for c in range(1, max_scan_cols + 1):
            t = _text(ws, r, c)
            if t and STUDENTS_RE.search(t):
                return r, c
    return None


def _find_data_start_row(ws, name_col, header_row, search_span=40):
    for r in range(header_row + 1, header_row + 1 + search_span):
        v = ws.cell(row=r, column=name_col).value
        if isinstance(v, str) and v.strip() and ("," in v or len(v.strip().split()) >= 2):
            return r
    return header_row + 4  # fallback guess


def parse_sheet(ws, sample_rows=3):
    anchor = _find_header_anchor(ws)
    if not anchor:
        return None
    header_row, name_col = anchor
    data_start = _find_data_start_row(ws, name_col, header_row)
    subheader_rows = list(range(header_row + 1, data_start))
    if not subheader_rows:
        return None

    # Step 1: section categories from row-`header_row` merges (e.g. "ATTENDANCE"
    # spanning K2:M2). Only single-row merges on the header row count -- multi-row
    # merges like the "STUDENTS" A2:C3 block are a different kind of header.
    # Blank-text merges are kept as bounded "unknown" spans -- a section title
    # box that exists but was never typed in -- so a later fallback can still
    # figure out what they are WITHOUT guessing past their real boundary.
    col_category = {}
    unknown_spans = []
    for mr in ws.merged_cells.ranges:
        if mr.min_row == header_row and mr.max_row == header_row and mr.max_col > mr.min_col:
            text = _text(ws, mr.min_row, mr.min_col)
            cat = _match_category(text)
            if cat:
                for c in range(mr.min_col, mr.max_col + 1):
                    col_category[c] = cat
            elif not text:
                unknown_spans.append((mr.min_col, mr.max_col))
    # Also catch a single-column category header sitting directly at header_row
    # (no merge), in case a sheet doesn't merge its section titles.
    for c in range(name_col + 1, ws.max_column + 1):
        if c not in col_category:
            cat = _match_category(_text(ws, header_row, c))
            if cat:
                col_category[c] = cat

    # Step 2: fallback for an attendance block with NO section title at all
    # (seen in the real Midterms sheet) -- if a column's sub-header holds an
    # actual date, it's an attendance column regardless of a missing label.
    for c in range(name_col + 1, ws.max_column + 1):
        if c in col_category:
            continue
        for r in subheader_rows:
            if isinstance(ws.cell(row=r, column=c).value, datetime):
                col_category[c] = "attendance"
                break

    # Step 2b: fallback for an UNLABELED raw block with no section title and
    # no dates (also seen in the real Midterms sheet: "Activities" section
    # merge exists but is blank text). These sections still give themselves
    # away because their raw columns are immediately followed by a
    # "<X> TOTAL" computed column -- e.g. "CA TOTAL" means the columns to its
    # left are Activities, "QUIZ TOTAL" means Quiz, "PT TOTAL" means PT.
    # The walk backward from the TOTAL column is bounded by the blank
    # section merge it falls inside (from Step 1), never past it -- so it
    # can't bleed into a neighboring, unrelated unlabeled section.
    total_prefix_to_category = {"ca": "activity", "quiz": "quiz", "pt": "pt", "performance task": "pt"}

    def _enclosing_unknown_span(col):
        for lo, hi in unknown_spans:
            if lo <= col <= hi:
                return lo, hi
        return None

    for c in range(name_col + 1, ws.max_column + 1):
        text = None
        for r in subheader_rows:
            t = _text(ws, r, c)
            if t:
                text = t
        m = re.match(r"^(ca|quiz|pt|performance\s*task)\s*total$", (text or "").strip().lower())
        if not m:
            continue
        category = total_prefix_to_category[re.sub(r"\s+", " ", m.group(1))]
        span = _enclosing_unknown_span(c - 1)
        left_bound = span[0] if span else name_col + 1
        left = c - 1
        while left >= left_bound and left not in col_category:
            col_category[left] = category
            left -= 1

    # Step 3: within each categorized column, raw input vs computed output.
    counters = {"attendance": 0, "activity": 0, "quiz": 0, "pt": 0, "exam": 0}
    slots = []
    for col in sorted(col_category):
        category = col_category[col]
        sub_text = None
        date_hint = None
        for r in subheader_rows:
            v = ws.cell(row=r, column=col).value
            if isinstance(v, datetime):
                date_hint = v.date().isoformat()
            elif isinstance(v, str) and v.strip():
                sub_text = v.strip()

        if _matches_any(sub_text, OUTPUT_ONLY_PATTERNS):
            continue  # CA TOTAL / % / GRADE etc. -- computed, not a slot

        sample = [ws.cell(row=r, column=col).value for r in range(data_start, data_start + sample_rows)]
        if any(_is_formula(v) for v in sample if v is not None):
            continue  # extra safety net: computed column we didn't catch by label

        counters[category] += 1
        if category == "attendance":
            label = date_hint or (sub_text or f"Attendance session {counters[category]}")
        elif sub_text:
            label = sub_text
        else:
            label = f"{category.capitalize()} {counters[category]}"

        slots.append({
            "category": category,
            "column": get_column_letter(col),
            "label": label,
            "date_hint": date_hint,
        })

    if not slots:
        return None
    return {
        "term_guess": _guess_term(ws.title),
        "header_row": header_row,
        "data_start_row": data_start,
        "name_column": get_column_letter(name_col),
        "slots": slots,
    }


def parse_workbook(path):
    """Returns {sheet_name: parse_result}, skipping sheets that don't look
    like a gradebook (no usable 'STUDENTS' header + raw slots found)."""
    wb = openpyxl.load_workbook(path, data_only=False)
    results = {}
    for sheet_name in wb.sheetnames:
        parsed = parse_sheet(wb[sheet_name])
        if parsed:
            results[sheet_name] = parsed
    return results
