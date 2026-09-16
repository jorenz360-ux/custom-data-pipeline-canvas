"""
Reproduces your original spreadsheet's exact formulas in Python, computed
fresh from raw Canvas scores every time -- nothing here is stored, so the
gradebook always reflects the current mapping + latest synced scores.

Midterm Grade  = ROUNDDOWN(CA*0.2 + Quiz*0.2 + Exam*0.6, 1)      -> PASS if <=3.0 else FAIL
Finals Grade   = ROUNDDOWN(CA*0.1 + Quiz*0.2 + PT*0.4 + Exam*0.3, 1)
Course Grade   = Midterm*0.5 + Finals*0.5

Where CA/Quiz/PT/Exam are each: grade_equivalent(ROUNDUP(total*100/max_points, 0))
Attendance is tracked (present-or-late = 1, else 0) but is NOT part of the
grade formula in your original template -- shown for reference only.
"""
import math
from datetime import datetime

from db import get_conn
from grading_scale import grade_equivalent

MIDTERM_CATEGORIES = ["activity", "quiz", "exam"]
FINALS_CATEGORIES_WITH_PT = ["activity", "quiz", "pt", "exam"]
FINALS_CATEGORIES_NO_PT = ["activity", "quiz", "exam"]
FINALS_CATEGORIES = FINALS_CATEGORIES_WITH_PT  # full set, for the Mapping page's category dropdown


def active_finals_categories(assignments=None):
    """Which Finals categories actually apply for this course: 'with PT'
    (CA 10% + Quiz 20% + PT 40% + Exam 30%) if any Finals+PT assignment is
    mapped and included, otherwise 'without PT' (CA 20% + Quiz 20% + Exam
    60%, same weight shape as Midterm) -- matching the original
    spreadsheet's two different Finals sheets/formulas. Pass an already-
    fetched assignments list to avoid a second query."""
    if assignments is None:
        conn = get_conn()
        assignments = conn.execute(
            "SELECT term, category FROM assignments WHERE included = 1"
        ).fetchall()
        conn.close()
    uses_pt = any(a["term"] == "finals" and a["category"] == "pt" for a in assignments)
    return FINALS_CATEGORIES_WITH_PT if uses_pt else FINALS_CATEGORIES_NO_PT


def _roundup(x, digits=0):
    factor = 10**digits
    return math.ceil(x * factor - 1e-9) / factor


def _rounddown(x, digits=0):
    factor = 10**digits
    return math.floor(x * factor + 1e-9) / factor


def _fmt_num(x):
    """Rounds a raw score/points value to the nearest whole number for
    display. Uses standard round-half-up (not Python's round-half-to-even),
    so e.g. 67.5 always rounds up to 68 -- the behavior instructors and
    students expect from a grade, matching the ROUNDUP/ROUNDDOWN convention
    used elsewhere in this file."""
    return str(math.floor(x + 0.5))


def _category_stats(scores_by_assignment, assignments):
    """assignments: list of assignment rows for one term+category.
    Returns (total, max_points, percent, grade) -- any of the last three can
    be None if no assignments are mapped yet."""
    if not assignments:
        return 0.0, 0.0, None, None
    total = sum((scores_by_assignment.get(a["id"]) or 0) for a in assignments)
    max_points = sum((a["points_possible"] or 0) for a in assignments)
    if max_points <= 0:
        return total, max_points, None, None
    percent = _roundup(total * 100 / max_points, 0)
    return total, max_points, percent, grade_equivalent(percent)


def compute_grades():
    conn = get_conn()
    students = conn.execute("SELECT * FROM students ORDER BY name").fetchall()
    assignments = conn.execute(
        "SELECT * FROM assignments WHERE included = 1 AND term IS NOT NULL AND category IS NOT NULL "
        "ORDER BY due_at, name"
    ).fetchall()
    all_scores = conn.execute("SELECT * FROM scores").fetchall()
    conn.close()

    scores_by_student = {}
    for row in all_scores:
        scores_by_student.setdefault(row["student_id"], {})[row["assignment_id"]] = row["score"]

    by_term_category = {}
    attendance_by_term = {"midterm": [], "finals": []}
    for a in assignments:
        if a["category"] == "attendance":
            attendance_by_term.setdefault(a["term"], []).append(a)
        else:
            by_term_category.setdefault((a["term"], a["category"]), []).append(a)

    finals_categories = active_finals_categories(assignments)
    finals_uses_pt = "pt" in finals_categories

    results = []
    for s in students:
        scores = scores_by_student.get(s["id"], {})
        row = {"student_id": s["id"], "name": s["name"], "categories": {}}

        term_grades = {}
        for term, categories in (("midterm", MIDTERM_CATEGORIES), ("finals", finals_categories)):
            cat_grades = {}
            complete = True
            for cat in categories:
                a_list = by_term_category.get((term, cat), [])
                total, max_points, percent, grade = _category_stats(scores, a_list)
                items = []
                for a in a_list:
                    score = scores.get(a["id"])
                    points = a["points_possible"]
                    if score is None:
                        display = "—"
                    elif points:
                        display = f"{_fmt_num(score)}/{_fmt_num(points)}"
                    else:
                        display = _fmt_num(score)
                    items.append({"score": score, "points_possible": points, "display": display})
                cat_grades[cat] = {
                    "total": total, "max": max_points, "percent": percent, "grade": grade,
                    "assignment_count": len(a_list), "scores": items,
                }
                if grade is None:
                    complete = False
            row["categories"][term] = cat_grades

            if complete:
                if term == "midterm":
                    weighted = (
                        cat_grades["activity"]["grade"] * 0.2
                        + cat_grades["quiz"]["grade"] * 0.2
                        + cat_grades["exam"]["grade"] * 0.6
                    )
                elif finals_uses_pt:
                    weighted = (
                        cat_grades["activity"]["grade"] * 0.1
                        + cat_grades["quiz"]["grade"] * 0.2
                        + cat_grades["pt"]["grade"] * 0.4
                        + cat_grades["exam"]["grade"] * 0.3
                    )
                else:
                    weighted = (
                        cat_grades["activity"]["grade"] * 0.2
                        + cat_grades["quiz"]["grade"] * 0.2
                        + cat_grades["exam"]["grade"] * 0.6
                    )
                term_grades[term] = _rounddown(weighted, 1)
            else:
                term_grades[term] = None

            # Attendance: tracked, not graded. present-or-late = 1, else 0
            # (including "no record yet"), matching the rule you specified.
            att_list = attendance_by_term.get(term, [])
            sessions = []
            present = 0
            for a in att_list:
                mark = 1 if (scores.get(a["id"]) or 0) > 0 else 0
                present += mark
                sessions.append({"name": a["name"], "mark": mark})
            row.setdefault("attendance", {})[term] = {
                "present": present, "total": len(att_list), "sessions": sessions,
            }

        row["midterm_grade"] = term_grades.get("midterm")
        row["finals_grade"] = term_grades.get("finals")
        if row["midterm_grade"] is not None and row["finals_grade"] is not None:
            row["course_grade"] = round(row["midterm_grade"] * 0.5 + row["finals_grade"] * 0.5, 2)
        else:
            row["course_grade"] = None
        row["midterm_status"] = (
            None if row["midterm_grade"] is None
            else ("PASS" if row["midterm_grade"] <= 3.0 else "FAIL")
        )

        results.append(row)

    return results


def category_assignment_names():
    """{"midterm": {cat: [assignment names, in column order]}, "finals": {...}} --
    the same order compute_grades() uses for each category's `items` list, so a
    caller can zip names with items to build per-assignment columns."""
    conn = get_conn()
    assignments = conn.execute(
        "SELECT * FROM assignments WHERE included = 1 AND term IS NOT NULL AND category IS NOT NULL "
        "ORDER BY due_at, name"
    ).fetchall()
    conn.close()
    headers = {
        "midterm": {c: [] for c in MIDTERM_CATEGORIES + ["attendance"]},
        "finals": {c: [] for c in FINALS_CATEGORIES + ["attendance"]},
    }
    for a in assignments:
        if a["term"] in headers and a["category"] in headers[a["term"]]:
            if a["category"] == "attendance":
                label = _attendance_date_label(a["due_at"]) or a["name"]
            else:
                label = a["name"]
            headers[a["term"]][a["category"]].append(label)
    return headers


def _attendance_date_label(due_at):
    """Formats a due date as a short column header (e.g. 'Sep 14') for
    per-day attendance assignments. Returns None if there's no due date,
    so the caller falls back to the assignment's name."""
    if not due_at:
        return None
    try:
        return datetime.fromisoformat(due_at.replace("Z", "+00:00")).strftime("%b %d")
    except ValueError:
        return None


def mapping_summary():
    """For the Mapping page: counts per term/category, so the instructor can
    see at a glance what still needs assignments."""
    conn = get_conn()
    rows = conn.execute(
        "SELECT term, category, COUNT(*) as n FROM assignments "
        "WHERE included = 1 AND term IS NOT NULL AND category IS NOT NULL "
        "GROUP BY term, category"
    ).fetchall()
    conn.close()
    summary = {}
    for r in rows:
        summary.setdefault(r["term"], {})[r["category"]] = r["n"]
    return summary
