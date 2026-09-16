"""
Imports a Canvas Roll Call "Attendance Report" CSV (Student ID, Student Name,
Class Date, Attendance, Timestamp columns -- extra columns are ignored) into
the database as per-date attendance "assignments".

Roll Call Attendance itself only ever exposes one aggregate percentage column
through Canvas's regular gradebook/API -- the day-by-day records live inside
the separate Roll Call LTI tool and only come out via this manually-exported
report. Each unique Class Date in the file becomes a synthetic assignment (a
negative canvas_assignment_id derived from the date, since real Canvas IDs
are always positive -- this also makes re-importing the same date update it
in place instead of creating a duplicate) tagged category='attendance' and
whichever term the user picked for this upload, so it flows through the same
grading/column-header code as any real Canvas assignment.
"""
import csv
import io
from datetime import datetime

from db import get_conn

PRESENT_STATUSES = {"present", "late", "excused"}
DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y")


def _synthetic_assignment_id(d):
    return -int(d.strftime("%Y%m%d"))


def _parse_date(s):
    s = (s or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def import_report(file_bytes, term):
    text = file_bytes.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    fieldnames = {(name or "").strip().lower(): name for name in (reader.fieldnames or [])}
    required = ["student id", "class date", "attendance"]
    missing = [c for c in required if c not in fieldnames]
    if missing:
        raise ValueError(f"missing column(s): {', '.join(missing)}")

    conn = get_conn()
    student_by_canvas_id = {
        r["canvas_user_id"]: r["id"] for r in conn.execute("SELECT id, canvas_user_id FROM students")
    }

    dates_seen = set()
    local_aid_by_synthetic = {}
    written = 0
    unmatched_students = set()
    bad_dates = 0

    for row in reader:
        d = _parse_date(row.get(fieldnames["class date"]))
        if d is None:
            bad_dates += 1
            continue
        try:
            canvas_uid = int((row.get(fieldnames["student id"]) or "").strip())
        except ValueError:
            continue
        status = (row.get(fieldnames["attendance"]) or "").strip().lower()

        aid = _synthetic_assignment_id(d)
        if aid not in dates_seen:
            dates_seen.add(aid)
            existing = conn.execute(
                "SELECT id FROM assignments WHERE canvas_assignment_id = ?", (aid,)
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE assignments SET term=?, category='attendance', included=1, "
                    "due_at=?, points_possible=1 WHERE canvas_assignment_id=?",
                    (term, d.isoformat(), aid),
                )
                local_aid_by_synthetic[aid] = existing["id"]
            else:
                cur = conn.execute(
                    "INSERT INTO assignments (canvas_assignment_id, name, points_possible, "
                    "due_at, canvas_group_name, term, category, included) "
                    "VALUES (?, 'Attendance', 1, ?, 'Attendance Report Import', ?, 'attendance', 1)",
                    (aid, d.isoformat(), term),
                )
                local_aid_by_synthetic[aid] = cur.lastrowid

        local_sid = student_by_canvas_id.get(canvas_uid)
        if local_sid is None:
            unmatched_students.add(canvas_uid)
            continue

        score = 1 if status in PRESENT_STATUSES else 0
        conn.execute(
            "INSERT INTO scores (student_id, assignment_id, score) VALUES (?, ?, ?) "
            "ON CONFLICT(student_id, assignment_id) DO UPDATE SET score = excluded.score",
            (local_sid, local_aid_by_synthetic[aid], score),
        )
        written += 1

    conn.commit()
    conn.close()
    return {
        "dates": len(dates_seen),
        "scores_written": written,
        "unmatched_students": len(unmatched_students),
        "bad_dates": bad_dates,
    }
