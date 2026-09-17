#!/usr/bin/env python3
"""
Initialize the concours_results SQLite database.

Safe to run multiple times: all tables use CREATE TABLE IF NOT EXISTS,
so existing data is never touched if the db already exists.

Every table has created_at and updated_at timestamp columns. SQLite has no
built-in "ON UPDATE CURRENT_TIMESTAMP" (unlike MySQL/Postgres), so
updated_at is kept current via an AFTER UPDATE trigger per table rather
than a column default (a default only fires on INSERT).

Usage:
    python init_db.py
"""

import sqlite3
from pathlib import Path
from config import DB_PATH

TABLES = [
    "banks",
    "schools",
    "classrooms",
    "students",
    "classroom_students",
    "admissions",
    "exams",
    "exam_results",
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS banks (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS schools (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT NOT NULL UNIQUE,
    bank_id    INTEGER NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (bank_id) REFERENCES banks (id)
);

CREATE TABLE IF NOT EXISTS classrooms (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    year       INTEGER NOT NULL UNIQUE,
    branch     TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS students (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    first_name TEXT,
    last_name  TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (first_name, last_name)
);

CREATE TABLE IF NOT EXISTS classroom_students (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    repeating    BOOLEAN NOT NULL DEFAULT 0,
    classroom_id INTEGER NOT NULL,
    student_id   INTEGER NOT NULL,
    created_at   TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at   TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (classroom_id) REFERENCES classrooms (id),
    FOREIGN KEY (student_id) REFERENCES students (id),
    UNIQUE (classroom_id, student_id)
);

CREATE TABLE IF NOT EXISTS admissions (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    classroom_student_id  INTEGER NOT NULL,
    school_id             INTEGER NOT NULL,
    written_status        TEXT,
    oral_status           TEXT,
    rank                  INTEGER,
    average               REAL,
    written_average       REAL,
    oral_average          REAL,
    total_points          REAL,
    written_points        REAL,
    oral_points           REAL,
    created_at            TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at            TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (classroom_student_id) REFERENCES classroom_students (id),
    FOREIGN KEY (school_id) REFERENCES schools (id),
    UNIQUE (classroom_student_id, school_id)
);

CREATE INDEX IF NOT EXISTS idx_admissions_rank ON admissions (rank);

CREATE TABLE IF NOT EXISTS exams (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    bank_id    INTEGER NOT NULL,
    name       TEXT NOT NULL,
    topic      TEXT,
    format     TEXT,
    weight     REAL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (bank_id) REFERENCES banks (id),
    UNIQUE (bank_id, name, format)
);

CREATE TABLE IF NOT EXISTS exam_results (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    classroom_student_id  INTEGER NOT NULL,
    exam_id               INTEGER NOT NULL,
    points                REAL,
    created_at            TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at            TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (classroom_student_id) REFERENCES classroom_students (id),
    FOREIGN KEY (exam_id) REFERENCES exams (id),
    UNIQUE (classroom_student_id, exam_id)
);
"""

TRIGGER_TEMPLATE = """
CREATE TRIGGER IF NOT EXISTS trg_{table}_updated_at
AFTER UPDATE ON {table}
FOR EACH ROW
BEGIN
    UPDATE {table} SET updated_at = datetime('now') WHERE id = NEW.id;
END;
"""


def init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    try:
        with conn:
            conn.executescript(SCHEMA)
            for table in TABLES:
                conn.executescript(TRIGGER_TEMPLATE.format(table=table))
    finally:
        conn.close()
    print(f"Database ready at {DB_PATH}")


if __name__ == "__main__":
    init_db()
