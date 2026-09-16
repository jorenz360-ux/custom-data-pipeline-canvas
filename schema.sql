CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS students (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    canvas_user_id  INTEGER UNIQUE NOT NULL,
    name            TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS assignments (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    canvas_assignment_id  INTEGER UNIQUE NOT NULL,
    name                  TEXT NOT NULL,
    points_possible       REAL,
    due_at                TEXT,
    canvas_group_name     TEXT,
    -- instructor-assigned mapping, all NULL until set on the Mapping page
    term                  TEXT,   -- 'midterm' | 'finals'
    category              TEXT,   -- 'attendance' | 'activity' | 'quiz' | 'pt' | 'exam'
    included              INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS scores (
    student_id    INTEGER NOT NULL REFERENCES students(id),
    assignment_id INTEGER NOT NULL REFERENCES assignments(id),
    score         REAL,     -- raw Canvas score, NULL if ungraded
    PRIMARY KEY (student_id, assignment_id)
);

CREATE TABLE IF NOT EXISTS sync_log (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    ran_at     TEXT NOT NULL,
    summary    TEXT NOT NULL
);
