"""
Talks to Canvas and writes into the local SQLite DB. Three operations:

- sync_assignments: refresh the assignment list from Canvas WITHOUT losing
  the term/category/included mapping you've already set on the Mapping page
  (matched by canvas_assignment_id, which never changes).
- sync_roster: refresh the student list.
- sync_scores: pull submissions for every assignment you've marked
  "included" and store the raw scores.
"""
from db import get_conn


def sync_assignments(client, course_id):
    conn = get_conn()
    groups = client.list_assignment_groups(course_id)
    added, updated = 0, 0
    for g in groups:
        group_name = g.get("name", "")
        for a in g.get("assignments", []):
            existing = conn.execute(
                "SELECT id FROM assignments WHERE canvas_assignment_id = ?",
                (a["id"],),
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE assignments SET name=?, points_possible=?, due_at=?, "
                    "canvas_group_name=? WHERE canvas_assignment_id=?",
                    (a["name"], a.get("points_possible"), a.get("due_at"), group_name, a["id"]),
                )
                updated += 1
            else:
                conn.execute(
                    "INSERT INTO assignments "
                    "(canvas_assignment_id, name, points_possible, due_at, canvas_group_name, "
                    " term, category, included) VALUES (?, ?, ?, ?, ?, NULL, NULL, 0)",
                    (a["id"], a["name"], a.get("points_possible"), a.get("due_at"), group_name),
                )
                added += 1
    conn.commit()
    conn.close()
    return {"added": added, "updated": updated}


def sync_roster(client, course_id):
    conn = get_conn()
    students = client.list_students(course_id)
    added, updated = 0, 0
    for s in students:
        name = s.get("sortable_name") or s.get("name")
        existing = conn.execute(
            "SELECT id FROM students WHERE canvas_user_id = ?", (s["id"],)
        ).fetchone()
        if existing:
            conn.execute("UPDATE students SET name=? WHERE canvas_user_id=?", (name, s["id"]))
            updated += 1
        else:
            conn.execute(
                "INSERT INTO students (canvas_user_id, name) VALUES (?, ?)", (s["id"], name)
            )
            added += 1
    conn.commit()
    conn.close()
    return {"added": added, "updated": updated}


def sync_scores(client, course_id):
    conn = get_conn()
    included = conn.execute(
        "SELECT id, canvas_assignment_id, name FROM assignments WHERE included = 1"
    ).fetchall()
    if not included:
        conn.close()
        return {"assignments": 0, "scores_written": 0, "failed": []}

    canvas_ids = [row["canvas_assignment_id"] for row in included]
    id_by_canvas_id = {row["canvas_assignment_id"]: row["id"] for row in included}
    name_by_canvas_id = {row["canvas_assignment_id"]: row["name"] for row in included}
    subs_by_assignment, errors = client.list_submissions_for_assignments(course_id, canvas_ids)

    student_rows = conn.execute("SELECT id, canvas_user_id FROM students").fetchall()
    local_student_id_by_canvas_uid = {r["canvas_user_id"]: r["id"] for r in student_rows}

    written = 0
    for canvas_aid, subs in subs_by_assignment.items():
        if canvas_aid in errors:
            continue  # fetch failed for this assignment -- leave its scores untouched
        local_aid = id_by_canvas_id[canvas_aid]
        for canvas_uid, sub in subs.items():
            local_sid = local_student_id_by_canvas_uid.get(canvas_uid)
            if local_sid is None:
                continue  # student not in roster sync yet (run Sync Roster first)
            score = sub.get("score")
            conn.execute(
                "INSERT INTO scores (student_id, assignment_id, score) VALUES (?, ?, ?) "
                "ON CONFLICT(student_id, assignment_id) DO UPDATE SET score = excluded.score",
                (local_sid, local_aid, score),
            )
            written += 1
    conn.commit()
    conn.close()
    failed = [
        {"id": aid, "name": name_by_canvas_id[aid], "error": err}
        for aid, err in errors.items()
    ]
    return {
        "assignments": len(canvas_ids) - len(failed),
        "scores_written": written,
        "failed": failed,
    }
