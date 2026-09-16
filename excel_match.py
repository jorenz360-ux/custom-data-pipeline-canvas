"""
Matches slots detected in an existing Excel gradebook (excel_importer.py)
against Canvas assignments already synced into the local database, so the
Mapping page can be pre-filled instead of tagged one-by-one from scratch.

Nothing here writes to the database -- it returns suggestions for the
Import preview page, which the instructor confirms (and can override) before
anything actually changes an assignment's term/category.
"""
from datetime import date, datetime

from db import get_conn

KEYWORD_BY_CATEGORY = {
    "attendance": ["attend"],
    "activity": ["activ", "class standing", "ca "],
    "quiz": ["quiz"],
    "pt": ["performance task", "pt"],
    "exam": ["exam"],
}


def _parse_date(s):
    if not s:
        return None
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def _due_date(a):
    if not a["due_at"]:
        return None
    try:
        return datetime.fromisoformat(a["due_at"].replace("Z", "+00:00")).date()
    except ValueError:
        return None


def _keyword_score(assignment_name, group_name, category):
    text = f"{assignment_name} {group_name or ''}".lower()
    return sum(1 for kw in KEYWORD_BY_CATEGORY.get(category, []) if kw in text)


def suggest_matches(parsed_sheets):
    """parsed_sheets: {sheet_name: parse_result} from excel_importer.parse_workbook.
    Returns a flat list of proposed rows, one per detected slot:
        {sheet, term, category, label, date_hint, slot_key,
         suggested_assignment_id, suggested_assignment_label, candidates: [...]}
    `candidates` is every unassigned-or-any Canvas assignment, for the dropdown.
    """
    conn = get_conn()
    assignments = conn.execute("SELECT * FROM assignments ORDER BY due_at, name").fetchall()
    conn.close()

    candidates = [
        {"id": a["id"], "label": f"{a['name']} ({a['points_possible'] or '?'} pts, due {(a['due_at'] or 'no date')[:10]})",
         "due_date": _due_date(a), "name": a["name"], "group_name": a["canvas_group_name"]}
        for a in assignments
    ]

    used_ids = set()
    proposals = []

    for sheet_name, parsed in parsed_sheets.items():
        term = parsed["term_guess"]
        for slot in parsed["slots"]:
            category = slot["category"]
            best = None
            best_score = -1

            if category == "attendance" and slot["date_hint"]:
                target = _parse_date(slot["date_hint"])
                for c in candidates:
                    if c["id"] in used_ids or c["due_date"] is None or target is None:
                        continue
                    # Require the attendance keyword too -- a close due date
                    # alone isn't enough (a Performance Task due the same
                    # week as an attendance date is not an attendance record).
                    if _keyword_score(c["name"], c["group_name"], "attendance") <= 0:
                        continue
                    delta = abs((c["due_date"] - target).days)
                    if delta <= 3:
                        score = 100 - delta  # closer date wins
                        if score > best_score:
                            best, best_score = c, score
            else:
                opposite_term_word = {"midterm": "final", "finals": "midterm"}.get(term)
                for c in candidates:
                    if c["id"] in used_ids:
                        continue
                    score = _keyword_score(c["name"], c["group_name"], category)
                    if score <= 0:
                        continue
                    # Don't suggest a "Final Exam"-named item for a Midterm
                    # slot or vice versa, even if it matches the category
                    # keyword -- term mismatch is a strong negative signal.
                    if opposite_term_word and opposite_term_word in c["name"].lower():
                        score -= 10
                    if score > best_score:
                        best, best_score = c, score
                if best_score <= 0:
                    best = None  # no (non-conflicting) keyword match -- don't guess randomly

            if best:
                used_ids.add(best["id"])

            proposals.append({
                "sheet": sheet_name,
                "term": term,
                "category": category,
                "label": slot["label"],
                "date_hint": slot["date_hint"],
                "slot_key": f"{sheet_name}|{slot['column']}",
                "suggested_assignment_id": best["id"] if best else None,
                "suggested_assignment_label": best["label"] if best else None,
                "candidates": candidates,
            })

    return proposals
