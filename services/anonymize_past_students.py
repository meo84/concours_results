"""
anonymize_past_students.py

Anonymizes students who were enrolled in a classroom before a given year,
and who are NOT enrolled in the classroom for that given year, by
setting their first_name and last_name to NULL in the concours_results
database (see init_db.py).

Usage:
    python anonymize_past_students.py 2026

Can also be imported and called as:
    from anonymize_past_students import anonymize_past_students
    anonymize_past_students(2026)
"""

import argparse
import sqlite3

from services import db_utils


def _find_students_to_anonymize(conn, before):
    """
    Student ids that were in a classroom for some year < `before`,
    and are NOT in the classroom for year == `before`.
    """
    rows = conn.execute(
        """
        SELECT DISTINCT s.id
        FROM students s
        JOIN classroom_students cs ON cs.student_id = s.id
        JOIN classrooms c ON c.id = cs.classroom_id
        WHERE c.year < ?
          AND s.id NOT IN (
              SELECT cs2.student_id
              FROM classroom_students cs2
              JOIN classrooms c2 ON c2.id = cs2.classroom_id
              WHERE c2.year >= ?
          )
        """,
        (before, before),
    ).fetchall()
    return [r[0] for r in rows]


def anonymize_past_students(before: int) -> int | None:
    """
    Set first_name and last_name to NULL for students who were in a
    classroom from a year before `before`, and who are not in the
    classroom for `before`.

    Returns the number of students anonymized.
    """
    conn = db_utils.get_connection()
    try:
        with conn:
            student_ids = _find_students_to_anonymize(conn, before)
            if student_ids:
                placeholders = ",".join("?" for _ in student_ids)
                conn.execute(
                    f"UPDATE students SET first_name = NULL, last_name = NULL "
                    f"WHERE id IN ({placeholders})",
                    student_ids,
                )
            count = len(student_ids)
            print(f"Anonymized {count} student(s)")
            return count
    except sqlite3.Error as e:
        print(f"Warning: failed to anonymize students before {before}: {e}")
        return None
    finally:
        conn.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=(
            "Anonymize students who were in a classroom before a given year "
            "and are not in the classroom for that year."
        )
    )
    parser.add_argument("before", type=int, help="Reference year")
    args = parser.parse_args()

    anonymize_past_students(args.before)
