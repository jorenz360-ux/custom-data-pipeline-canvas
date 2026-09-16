import csv
import io
import os
import tempfile
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, flash, Response
from werkzeug.utils import secure_filename

from db import init_db, get_conn, get_setting, set_setting
from canvas_client import CanvasClient
import canvas_sync
import grading
import excel_importer
import excel_match
import attendance_report
import gradebook_export

app = Flask(__name__)
app.secret_key = "local-only-admin-page-no-real-secret-needed"

CATEGORY_LABELS = {
    "attendance": "Attendance",
    "activity": "Activity (CA)",
    "quiz": "Quiz",
    "pt": "Performance Task",
    "exam": "Exam",
}


def get_client():
    base_url = get_setting("canvas_base_url")
    token = get_setting("canvas_token")
    if not base_url or not token:
        return None
    return CanvasClient(base_url, token)


@app.context_processor
def inject_globals():
    return {
        "course_id": get_setting("course_id"),
        "course_name": get_setting("course_name"),
        "configured": bool(get_setting("canvas_base_url") and get_setting("canvas_token")),
    }


@app.route("/")
def index():
    if not get_setting("canvas_base_url") or not get_setting("canvas_token"):
        return redirect(url_for("setup"))
    if not get_setting("course_id"):
        return redirect(url_for("courses"))
    return redirect(url_for("gradebook"))


@app.route("/setup", methods=["GET", "POST"])
def setup():
    if request.method == "POST":
        set_setting("canvas_base_url", request.form["base_url"].strip().rstrip("/"))
        set_setting("canvas_token", request.form["token"].strip())
        client = get_client()
        try:
            profile = client.test_connection()
            flash(f"Connected as {profile.get('name', 'unknown user')}.", "success")
            return redirect(url_for("courses"))
        except Exception as e:
            flash(f"Could not connect: {e}", "error")
    return render_template(
        "setup.html",
        base_url=get_setting("canvas_base_url", ""),
        token=get_setting("canvas_token", ""),
    )


@app.route("/courses", methods=["GET", "POST"])
def courses():
    client = get_client()
    if not client:
        return redirect(url_for("setup"))

    if request.method == "POST":
        new_course_id = request.form["course_id"]
        old_course_id = get_setting("course_id")
        cleared = False
        if old_course_id and old_course_id != new_course_id:
            conn = get_conn()
            conn.execute("DELETE FROM scores")
            conn.execute("DELETE FROM assignments")
            conn.execute("DELETE FROM students")
            conn.commit()
            conn.close()
            cleared = True

        set_setting("course_id", new_course_id)
        set_setting("course_name", request.form["course_name"])
        msg = "Course selected. Now go sync assignments and roster."
        if cleared:
            msg = "Switched course -- cleared the previous course's roster/assignments/scores. " + msg
        flash(msg, "success")
        return redirect(url_for("mapping"))

    try:
        course_list = client.list_courses()
    except Exception as e:
        flash(f"Could not fetch courses: {e}", "error")
        course_list = []
    return render_template("courses.html", courses=course_list)


@app.route("/sync/<what>", methods=["POST"])
def sync(what):
    client = get_client()
    course_id = get_setting("course_id")
    if not client or not course_id:
        flash("Set up Canvas connection and pick a course first.", "error")
        return redirect(url_for("index"))

    try:
        if what == "assignments":
            result = canvas_sync.sync_assignments(client, course_id)
            msg = f"Assignments: {result['added']} added, {result['updated']} updated."
        elif what == "roster":
            result = canvas_sync.sync_roster(client, course_id)
            msg = f"Roster: {result['added']} added, {result['updated']} updated."
        elif what == "scores":
            result = canvas_sync.sync_scores(client, course_id)
            msg = f"Scores: pulled {result['assignments']} assignment(s), wrote {result['scores_written']} score(s)."
            if result["failed"]:
                details = "; ".join(f"\"{f['name']}\" ({f['error']})" for f in result["failed"])
                msg += f" {len(result['failed'])} assignment(s) failed and were skipped: {details}"
        else:
            flash("Unknown sync target.", "error")
            return redirect(url_for("mapping"))

        conn = get_conn()
        conn.execute(
            "INSERT INTO sync_log (ran_at, summary) VALUES (?, ?)",
            (datetime.now().isoformat(timespec="seconds"), msg),
        )
        conn.commit()
        conn.close()
        flash(msg, "success")
    except Exception as e:
        flash(f"Sync failed: {e}", "error")

    return redirect(request.referrer or url_for("mapping"))


@app.route("/mapping", methods=["GET"])
def mapping():
    conn = get_conn()
    assignments = conn.execute(
        "SELECT * FROM assignments ORDER BY canvas_group_name, due_at, name"
    ).fetchall()
    log = conn.execute(
        "SELECT * FROM sync_log ORDER BY id DESC LIMIT 5"
    ).fetchall()
    conn.close()

    groups = {}
    for a in assignments:
        groups.setdefault(a["canvas_group_name"] or "(no group)", []).append(a)

    return render_template(
        "mapping.html",
        groups=groups,
        category_labels=CATEGORY_LABELS,
        summary=grading.mapping_summary(),
        log=log,
        has_assignments=bool(assignments),
    )


@app.route("/mapping/save", methods=["POST"])
def mapping_save():
    conn = get_conn()
    assignments = conn.execute("SELECT id FROM assignments").fetchall()
    for a in assignments:
        aid = a["id"]
        term = request.form.get(f"term_{aid}") or None
        category = request.form.get(f"category_{aid}") or None
        included = 1 if request.form.get(f"included_{aid}") else 0
        conn.execute(
            "UPDATE assignments SET term=?, category=?, included=? WHERE id=?",
            (term, category, included, aid),
        )
    conn.commit()
    conn.close()
    flash("Mapping saved.", "success")
    return redirect(url_for("mapping"))


IMPORT_TMP_DIR = os.path.join(tempfile.gettempdir(), "canvas_gradebook_imports")
os.makedirs(IMPORT_TMP_DIR, exist_ok=True)


@app.route("/import", methods=["GET", "POST"])
def import_excel():
    if request.method == "GET":
        return render_template("import.html")

    file = request.files.get("excel_file")
    if not file or not file.filename:
        flash("Choose an .xlsx file first.", "error")
        return redirect(url_for("import_excel"))

    token = secure_filename(f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
    saved_path = os.path.join(IMPORT_TMP_DIR, token)
    file.save(saved_path)

    try:
        parsed_sheets = excel_importer.parse_workbook(saved_path)
    except Exception as e:
        flash(f"Could not read that file: {e}", "error")
        return redirect(url_for("import_excel"))
    finally:
        os.remove(saved_path)

    if not parsed_sheets:
        flash(
            "Couldn't find any recognizable gradebook structure in that file "
            "(no 'STUDENTS' header found, or no raw score columns detected).",
            "error",
        )
        return redirect(url_for("import_excel"))

    proposals = excel_match.suggest_matches(parsed_sheets)
    if not proposals:
        flash("Parsed the file but found no usable slots.", "error")
        return redirect(url_for("import_excel"))

    import json
    review_token = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}.json"
    with open(os.path.join(IMPORT_TMP_DIR, review_token), "w") as f:
        json.dump(proposals, f, default=str)

    return render_template(
        "import_preview.html",
        proposals=proposals,
        review_token=review_token,
        category_labels=CATEGORY_LABELS,
    )


@app.route("/import/apply", methods=["POST"])
def import_apply():
    import json
    review_token = request.form.get("review_token", "")
    path = os.path.join(IMPORT_TMP_DIR, secure_filename(review_token))
    if not os.path.isfile(path):
        flash("That import preview expired -- please upload the file again.", "error")
        return redirect(url_for("import_excel"))

    with open(path) as f:
        proposals = json.load(f)
    os.remove(path)

    conn = get_conn()
    applied, skipped = 0, 0
    for p in proposals:
        key = p["slot_key"]
        assignment_id = request.form.get(f"assignment_{key}")
        term = request.form.get(f"term_{key}") or None
        category = request.form.get(f"category_{key}") or None
        if not assignment_id or not term or not category:
            skipped += 1
            continue
        conn.execute(
            "UPDATE assignments SET term=?, category=?, included=1 WHERE id=?",
            (term, category, int(assignment_id)),
        )
        applied += 1
    conn.commit()
    conn.close()

    flash(f"Applied {applied} slot mapping(s) from the Excel import, skipped {skipped} unmatched.", "success")
    return redirect(url_for("mapping"))


@app.route("/attendance/import", methods=["POST"])
def import_attendance_report():
    file = request.files.get("attendance_csv")
    term = request.form.get("term")
    if not file or not file.filename:
        flash("Choose a CSV file first.", "error")
        return redirect(url_for("mapping"))
    if term not in ("midterm", "finals"):
        flash("Pick a term for this attendance report.", "error")
        return redirect(url_for("mapping"))

    try:
        result = attendance_report.import_report(file.read(), term)
    except ValueError as e:
        flash(f"Could not read that file -- {e}.", "error")
        return redirect(url_for("mapping"))

    msg = f"Attendance import: {result['dates']} date(s), {result['scores_written']} record(s) written."
    problems = []
    if result["unmatched_students"]:
        problems.append(f"{result['unmatched_students']} student ID(s) not found in your roster (run Sync Roster first)")
    if result["bad_dates"]:
        problems.append(f"{result['bad_dates']} row(s) had an unreadable Class Date and were skipped")
    if problems:
        msg += " " + "; ".join(problems) + "."
    flash(msg, "error" if problems else "success")
    return redirect(url_for("mapping"))


@app.route("/gradebook")
def gradebook():
    rows = grading.compute_grades()
    column_headers = grading.category_assignment_names()
    finals_categories = grading.active_finals_categories()
    midterm_colspan = (
        len(column_headers["midterm"]["attendance"])
        + sum(len(column_headers["midterm"][c]) + 1 for c in grading.MIDTERM_CATEGORIES)
    )
    finals_colspan = (
        len(column_headers["finals"]["attendance"])
        + sum(len(column_headers["finals"][c]) + 1 for c in finals_categories)
    )
    return render_template(
        "gradebook.html", rows=rows,
        midterm_categories=grading.MIDTERM_CATEGORIES,
        finals_categories=finals_categories,
        category_labels=CATEGORY_LABELS,
        column_headers=column_headers,
        midterm_colspan=midterm_colspan,
        finals_colspan=finals_colspan,
    )


@app.route("/gradebook/export.csv")
def export_csv():
    rows = grading.compute_grades()
    column_headers = grading.category_assignment_names()
    finals_categories = grading.active_finals_categories()

    header = ["Name"]
    for term, categories in (("midterm", grading.MIDTERM_CATEGORIES), ("finals", finals_categories)):
        term_label = term.capitalize()
        for name in column_headers[term]["attendance"]:
            header.append(f"{term_label} Attendance: {name}")
        for cat in categories:
            cat_label = CATEGORY_LABELS[cat]
            for name in column_headers[term][cat]:
                header.append(f"{term_label} {cat_label}: {name}")
            header.append(f"{term_label} {cat_label} Grade")
        if term == "midterm":
            header += ["Midterm Grade", "Midterm Status"]
        else:
            header += ["Finals Grade", "Course Grade"]
    header += ["Midterm Attendance (present/total)", "Finals Attendance (present/total)"]

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(header)
    for r in rows:
        line = [r["name"]]
        for term, categories in (("midterm", grading.MIDTERM_CATEGORIES), ("finals", finals_categories)):
            cats = r["categories"][term]
            for session in r["attendance"][term]["sessions"]:
                line.append(session["mark"])
            for cat in categories:
                for item in cats[cat]["scores"]:
                    line.append(item["display"])
                line.append(cats[cat]["grade"])
            if term == "midterm":
                line += [r["midterm_grade"], r["midterm_status"]]
            else:
                line += [r["finals_grade"], r["course_grade"]]
        ma = r["attendance"]["midterm"]
        fa = r["attendance"]["finals"]
        line += [f"{ma['present']}/{ma['total']}", f"{fa['present']}/{fa['total']}"]
        writer.writerow(line)

    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=gradebook_export.csv"},
    )


@app.route("/gradebook/export.xlsx")
def export_xlsx():
    course_name = get_setting("course_name") or ""
    data = gradebook_export.build_workbook(CATEGORY_LABELS, course_name)
    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=gradebook_export.xlsx"},
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=True, port=5055)
