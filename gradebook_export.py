"""
Builds an .xlsx download that mirrors the on-screen Gradebook page: one sheet
per term (Midterm, Finals) plus a "Grading" lookup sheet, each with a grouped
header (category -> one column per attendance date/assignment -> Total
Points -> % -> Grade) and the same attendance/status coloring you see in the
browser.

Unlike the plain-value CSV export, every *computed* cell here is a live Excel
formula (Total Points = SUM of the raw score cells, % = ROUNDUP against a Max
Points reference row, Grade = VLOOKUP into the Grading sheet, Midterm/Finals/
Course Grade = the weighted formulas from grading.py, Status = PASS/FAIL) --
so hand-editing a raw score in Excel (to fix something Canvas never captured)
recalculates everything downstream automatically. Raw score cells and
attendance 1/0 marks are the only plain values, since they're inputs, not
calculations.
"""
import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

import grading
import grading_scale

HEADER_FILL = PatternFill("solid", fgColor="FFFF00")
HEADER_FONT = Font(color="000000", bold=True)
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
DATE_FILL = PatternFill("solid", fgColor="0070C0")
DATE_FONT = Font(color="FFFFFF", bold=True)
REF_FONT = Font(italic=True, color="6B7280")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

TITLE_FONT = Font(bold=True, size=14)
TITLE_ALIGN = Alignment(horizontal="center", vertical="center")

_THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

PRESENT_FILL = PatternFill("solid", fgColor="1E8449")
PRESENT_FONT = Font(color="FFFFFF", bold=True)
ABSENT_FILL = PatternFill("solid", fgColor="C0392B")
ABSENT_FONT = Font(color="FFFFFF", bold=True)
PASS_FILL = PRESENT_FILL
FAIL_FILL = ABSENT_FILL

TITLE_ROW, GROUP_ROW, LABEL_ROW, MAXPTS_ROW, DATA_START_ROW = 1, 2, 3, 4, 5

MIDTERM_WEIGHTS = {"activity": 0.2, "quiz": 0.2, "exam": 0.6}
FINALS_WEIGHTS_WITH_PT = {"activity": 0.1, "quiz": 0.2, "pt": 0.4, "exam": 0.3}
FINALS_WEIGHTS_NO_PT = {"activity": 0.2, "quiz": 0.2, "exam": 0.6}

NON_ITEM_LABELS = ("Grade", "Total Points", "%")


def _term_column_plan(term_key, categories, category_labels, column_headers, trailing_labels):
    """Flat left-to-right column list for one term's sheet, as (subgroup,
    label) tuples, plus a `layout` dict recording each category's raw-score
    column range and its Total Points/%/Grade column indexes, and the
    trailing columns' indexes -- so formulas can be built from fixed
    positions instead of re-deriving them per row."""
    plan = [("", "Student")]
    layout = {"categories": {}, "trailing": {}}
    col = 2

    for name in column_headers[term_key]["attendance"]:
        plan.append(("Attendance", name))
        col += 1

    for cat in categories:
        label = category_labels[cat]
        names = column_headers[term_key][cat]
        raw_start = col if names else None
        for name in names:
            plan.append((label, name))
            col += 1
        raw_end = col - 1 if names else None

        total_col = col
        plan.append((label, "Total Points"))
        col += 1
        pct_col = col
        plan.append((label, "%"))
        col += 1
        grade_col = col
        plan.append((label, "Grade"))
        col += 1

        layout["categories"][cat] = {
            "raw_start": raw_start, "raw_end": raw_end,
            "total_col": total_col, "pct_col": pct_col, "grade_col": grade_col,
        }

    for label in trailing_labels:
        plan.append(("", label))
        layout["trailing"][label] = col
        col += 1

    return plan, layout


def _write_title(ws, text, n_cols):
    cell = ws.cell(row=TITLE_ROW, column=1, value=text)
    cell.font = TITLE_FONT
    cell.alignment = TITLE_ALIGN
    if n_cols > 1:
        ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=n_cols)


def _write_header(ws, plan, white_name_labels):
    n_cols = len(plan)

    for col_idx, (subgroup, label) in enumerate(plan, start=1):
        ws.cell(row=GROUP_ROW if not subgroup else LABEL_ROW, column=col_idx, value=label)

    col = 1
    subgroup_runs = []
    while col <= n_cols:
        key = plan[col - 1][0]
        start = col
        col += 1
        while col <= n_cols and plan[col - 1][0] == key and key:
            col += 1
        if key:
            ws.cell(row=GROUP_ROW, column=start, value=key)
        subgroup_runs.append((key, start, col - 1))

    for row in (GROUP_ROW, LABEL_ROW):
        for col_idx in range(1, n_cols + 1):
            subgroup, label = plan[col_idx - 1]
            is_item_name = (
                row == LABEL_ROW and subgroup in white_name_labels
                and label not in NON_ITEM_LABELS
            )
            is_attendance_date = row == LABEL_ROW and subgroup == "Attendance"
            cell = ws.cell(row=row, column=col_idx)
            if is_item_name:
                cell.fill, cell.font = WHITE_FILL, HEADER_FONT
            elif is_attendance_date:
                cell.fill, cell.font = DATE_FILL, DATE_FONT
            else:
                cell.fill, cell.font = HEADER_FILL, HEADER_FONT
            cell.alignment = CENTER
            cell.border = BORDER

    for key, start, end in subgroup_runs:
        if key and end > start:
            ws.merge_cells(start_row=GROUP_ROW, start_column=start, end_row=GROUP_ROW, end_column=end)
    for col_idx, (subgroup, label) in enumerate(plan, start=1):
        if not subgroup:
            ws.merge_cells(start_row=GROUP_ROW, start_column=col_idx, end_row=LABEL_ROW, end_column=col_idx)


def _write_max_points_row(ws, layout, categories, sample_term_data, n_cols):
    """Row 4: a reference row holding each raw column's points_possible and
    each category's total max -- the % formula's denominator, referenced by
    absolute cell address so it isn't hardcoded into the formula text."""
    ws.cell(row=MAXPTS_ROW, column=1, value="Max Points")
    for cat in categories:
        lay = layout["categories"][cat]
        catdata = sample_term_data[cat] if sample_term_data else None
        if lay["raw_start"] is not None and catdata:
            for i, item in enumerate(catdata["scores"]):
                ws.cell(row=MAXPTS_ROW, column=lay["raw_start"] + i, value=item["points_possible"])
        ws.cell(row=MAXPTS_ROW, column=lay["total_col"], value=catdata["max"] if catdata else 0)
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=MAXPTS_ROW, column=col_idx)
        cell.font = REF_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def _autosize(ws, rows, n_cols, data_start_row):
    name_len = max((len(row_data["name"]) for row_data in rows), default=10)
    ws.column_dimensions["A"].width = min(max(name_len + 2, 12), 30)
    for col_idx in range(2, n_cols + 1):
        label_len = len(str(ws.cell(row=LABEL_ROW, column=col_idx).value or ""))
        data_len = max(
            (len(str(ws.cell(row=r_idx, column=col_idx).value))
             for r_idx in range(data_start_row, ws.max_row + 1)
             if ws.cell(row=r_idx, column=col_idx).value is not None
             and not str(ws.cell(row=r_idx, column=col_idx).value).startswith("=")),
            default=0,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max(label_len, data_len) + 2, 6), 12)


def _weighted_grade_formula(layout, categories, weights, row):
    grade_cells = [f"{get_column_letter(layout['categories'][cat]['grade_col'])}{row}" for cat in categories]
    guard = "OR(" + ",".join(f'{c}=""' for c in grade_cells) + ")"
    body = "+".join(f"({c}*{weights[cat]})" for c, cat in zip(grade_cells, categories))
    return f'=IF({guard},"",ROUNDDOWN({body},1))'


def _build_term_sheet(ws, term_key, categories, rows, category_labels, column_headers, trailing,
                       title, midterm_layout=None):
    """trailing: list of (kind, header_label) appended after the category
    columns. kind is one of: 'midterm_grade_here' (this term's own Midterm
    Grade formula), 'midterm_grade_ref' (cross-sheet copy of the Midterm
    sheet's Midterm Grade, for the Finals sheet), 'finals_grade',
    'course_grade', 'status', 'attendance_summary'.
    midterm_layout: the Midterm sheet's layout dict, needed only when
    building the Finals sheet (for the cross-sheet Midterm Grade reference)."""
    plan, layout = _term_column_plan(term_key, categories, category_labels, column_headers,
                                      [label for _, label in trailing])
    n_cols = len(plan)
    _write_title(ws, title, n_cols)
    white_name_labels = {category_labels.get(k) for k in ("activity", "quiz", "exam")}
    _write_header(ws, plan, white_name_labels)

    sample_term_data = rows[0]["categories"][term_key] if rows else None
    _write_max_points_row(ws, layout, categories, sample_term_data, n_cols)

    finals_uses_pt = "pt" in categories
    weights = (MIDTERM_WEIGHTS if term_key == "midterm"
               else (FINALS_WEIGHTS_WITH_PT if finals_uses_pt else FINALS_WEIGHTS_NO_PT))

    r = DATA_START_ROW
    for row_data in rows:
        name_cell = ws.cell(row=r, column=1, value=row_data["name"])
        name_cell.border = BORDER

        for i, session in enumerate(row_data["attendance"][term_key]["sessions"]):
            col = 2 + i
            cell = ws.cell(row=r, column=col, value=session["mark"])
            cell.alignment = CENTER
            cell.border = BORDER
            cell.fill, cell.font = (PRESENT_FILL, PRESENT_FONT) if session["mark"] else (ABSENT_FILL, ABSENT_FONT)

        for cat in categories:
            lay = layout["categories"][cat]
            catdata = row_data["categories"][term_key][cat]
            total_letter = get_column_letter(lay["total_col"])
            pct_letter = get_column_letter(lay["pct_col"])

            if lay["raw_start"] is not None:
                for i, item in enumerate(catdata["scores"]):
                    ws.cell(row=r, column=lay["raw_start"] + i, value=item["score"]).border = BORDER
                rng = f"{get_column_letter(lay['raw_start'])}{r}:{get_column_letter(lay['raw_end'])}{r}"
                ws.cell(row=r, column=lay["total_col"], value=f"=SUM({rng})").border = BORDER
            else:
                ws.cell(row=r, column=lay["total_col"], value=0).border = BORDER

            max_ref = f"${total_letter}${MAXPTS_ROW}"
            pct_formula = f'=IF({max_ref}=0,"",ROUNDUP({total_letter}{r}*100/{max_ref},0))'
            ws.cell(row=r, column=lay["pct_col"], value=pct_formula).border = BORDER

            grade_formula = f'=IF({pct_letter}{r}="","",VLOOKUP({pct_letter}{r},Grading!$A$1:$B$101,2,TRUE))'
            ws.cell(row=r, column=lay["grade_col"], value=grade_formula).border = BORDER

        for kind, label in trailing:
            col = layout["trailing"][label]
            if kind == "midterm_grade_here":
                value = _weighted_grade_formula(layout, grading.MIDTERM_CATEGORIES, MIDTERM_WEIGHTS, r)
            elif kind == "midterm_grade_ref":
                mg_col = get_column_letter(midterm_layout["trailing"]["Midterm Grade"])
                value = f"=Midterm!{mg_col}{r}"
            elif kind == "finals_grade":
                value = _weighted_grade_formula(layout, categories, weights, r)
            elif kind == "course_grade":
                mg_col = get_column_letter(layout["trailing"]["Midterm Grade"])
                fg_col = get_column_letter(layout["trailing"]["Finals Grade"])
                value = f'=IF(OR({mg_col}{r}="",{fg_col}{r}=""),"",({mg_col}{r}*0.5)+({fg_col}{r}*0.5))'
            elif kind == "status":
                mg_col = get_column_letter(layout["trailing"]["Midterm Grade"])
                value = f'=IF({mg_col}{r}="","",IF({mg_col}{r}<=3,"PASS","FAIL"))'
            else:  # attendance_summary
                att = row_data["attendance"][term_key]
                value = f"{att['present']}/{att['total']}"
            cell = ws.cell(row=r, column=col, value=value)
            cell.alignment = CENTER
            cell.border = BORDER

        r += 1

    if "Status" in layout["trailing"]:
        status_letter = get_column_letter(layout["trailing"]["Status"])
        rng = f"{status_letter}{DATA_START_ROW}:{status_letter}{r - 1}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'], fill=PASS_FILL))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"FAIL"'], fill=FAIL_FILL))

    ws.freeze_panes = f"B{DATA_START_ROW}"
    _autosize(ws, rows, n_cols, DATA_START_ROW)
    return layout


def _write_grading_sheet(wb):
    ws = wb.create_sheet("Grading")
    for percent, grade in enumerate(grading_scale.TABLE):
        ws.cell(row=percent + 1, column=1, value=percent)
        ws.cell(row=percent + 1, column=2, value=grade)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.sheet_state = "hidden"


def build_workbook(category_labels, course_name=""):
    rows = grading.compute_grades()
    column_headers = grading.category_assignment_names()
    finals_categories = grading.active_finals_categories()
    title_prefix = f"{course_name} — " if course_name else ""

    wb = Workbook()
    ws_mid = wb.active
    ws_mid.title = "Midterm"
    midterm_layout = _build_term_sheet(
        ws_mid, "midterm", grading.MIDTERM_CATEGORIES, rows, category_labels, column_headers,
        trailing=[
            ("midterm_grade_here", "Midterm Grade"),
            ("status", "Status"),
            ("attendance_summary", "Midterm Attendance (present/total)"),
        ],
        title=f"{title_prefix}Midterm",
    )

    ws_fin = wb.create_sheet("Finals")
    _build_term_sheet(
        ws_fin, "finals", finals_categories, rows, category_labels, column_headers,
        trailing=[
            ("midterm_grade_ref", "Midterm Grade"),
            ("finals_grade", "Finals Grade"),
            ("course_grade", "Course Grade"),
            ("attendance_summary", "Finals Attendance (present/total)"),
        ],
        title=f"{title_prefix}Finals",
        midterm_layout=midterm_layout,
    )

    _write_grading_sheet(wb)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
