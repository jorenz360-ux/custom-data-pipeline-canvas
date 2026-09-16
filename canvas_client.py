"""Minimal Canvas LMS REST API client. Auth via Bearer token, pagination via
the Link header. Docs: https://canvas.instructure.com/doc/api/"""
import requests


class CanvasClient:
    def __init__(self, base_url: str, token: str):
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.session = requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {self.token}"})

    def _get_paginated(self, path: str, params: dict = None):
        url = f"{self.base_url}/api/v1{path}"
        results = []
        params = dict(params or {})
        params.setdefault("per_page", 100)
        while url:
            resp = self.session.get(url, params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            if isinstance(data, list):
                results.extend(data)
            else:
                results.append(data)
            params = None
            url = resp.links.get("next", {}).get("url")
        return results

    def test_connection(self):
        """Raises if the token/URL are bad; returns the profile on success."""
        resp = self.session.get(f"{self.base_url}/api/v1/users/self", timeout=15)
        resp.raise_for_status()
        return resp.json()

    def list_courses(self):
        return self._get_paginated("/courses", {"enrollment_state": "active"})

    def list_students(self, course_id):
        return self._get_paginated(
            f"/courses/{course_id}/users",
            {"enrollment_type[]": "student", "enrollment_state[]": "active"},
        )

    def list_assignment_groups(self, course_id):
        return self._get_paginated(
            f"/courses/{course_id}/assignment_groups", {"include[]": "assignments"}
        )

    def list_submissions_for_assignments(self, course_id, assignment_ids):
        """Returns (by_assignment, errors).

        by_assignment: {assignment_id: {user_id: submission_dict}}
        errors: {assignment_id: str} for any assignment whose request failed
            (e.g. 403 because that specific assignment uses moderated/anonymous
            grading, was deleted, etc.) -- fetched one assignment at a time so
            one bad assignment can't block scores for the rest.
        """
        by_assignment = {aid: {} for aid in assignment_ids}
        errors = {}
        for aid in assignment_ids:
            params = {"student_ids[]": "all", "assignment_ids[]": [str(aid)]}
            try:
                subs = self._get_paginated(f"/courses/{course_id}/students/submissions", params)
            except requests.HTTPError as e:
                errors[aid] = str(e)
                continue
            for sub in subs:
                uid = sub.get("user_id")
                by_assignment[aid][uid] = sub
        return by_assignment, errors
