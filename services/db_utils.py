"""
Shared helpers for the concours_results import scripts.

Keeping these here means each import_*.py script stays a thin, independently
runnable step (matching how you run them in sequence), while sharing the
same DB-access and get-or-create logic instead of duplicating it.
"""

import sqlite3
from collections.abc import Hashable, Iterator
from contextlib import contextmanager

from config import DB_PATH
from models.admission import ExamFormat
from models.classroom import Branch
from models.text import normalize_name


def get_connection(create: bool = False) -> sqlite3.Connection:
    """Open the database with foreign keys enforced and rows accessible by name.

    Raises FileNotFoundError if the database doesn't exist yet, unless `create`
    is set, in which case SQLite creates it.
    """
    if not create and not DB_PATH.exists():
        raise FileNotFoundError(
            f"Database not found at {DB_PATH}. Run init_db.py first."
        )
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


Cache = dict[Hashable, int]
"""Ids already found by name, keyed by (scoped) normalized name."""


@contextmanager
def transaction() -> Iterator[sqlite3.Connection]:
    """Open a connection, commit on success, roll back on error, then close."""
    conn = get_connection()
    try:
        with conn:
            yield conn
    finally:
        conn.close()


def _find_id(
    conn: sqlite3.Connection, table: str, where_sql: str, params: tuple = ()
) -> int | None:
    """Return the id of the first row of `table` matching the WHERE clause."""
    row = conn.execute(f"SELECT id FROM {table} {where_sql}", params).fetchone()
    return row[0] if row else None


def _get_or_create(
    conn: sqlite3.Connection,
    table: str,
    where_sql: str,
    params: tuple,
    insert_sql: str,
    insert_params: tuple,
) -> tuple[int, bool]:
    """Exact match, creating the row if none exists.

    Returns (row_id, created).
    """
    row_id = _find_id(conn, table, where_sql, params)
    if row_id is not None:
        return row_id, False

    cur = conn.execute(insert_sql, insert_params)
    return cur.lastrowid, True


def _find_id_loose(
    conn: sqlite3.Connection,
    table: str,
    name: str,
    cache: Cache,
    cache_key: Hashable,
    where_sql: str = "",
    params: tuple = (),
) -> int | None:
    """Loose match (accent/punctuation/case-insensitive) on name, optionally scoped.

    Found ids are remembered in `cache` under `cache_key`.
    """
    if cache_key in cache:
        return cache[cache_key]

    target = normalize_name(name)

    rows = conn.execute(f"SELECT id, name FROM {table} {where_sql}", params).fetchall()

    for row_id, existing_name in rows:
        if normalize_name(existing_name) == target:
            cache[cache_key] = row_id
            return row_id

    return None


def _get_or_create_loose(
    conn: sqlite3.Connection,
    table: str,
    name: str,
    cache: Cache,
    cache_key: Hashable,
    insert_sql: str,
    insert_params: tuple,
    where_sql: str = "",
    params: tuple = (),
) -> tuple[int, bool]:
    """Loose match on name, creating the row if none exists.

    Returns (row_id, created).
    """
    row_id = _find_id_loose(conn, table, name, cache, cache_key, where_sql, params)
    if row_id is not None:
        return row_id, False

    cur = conn.execute(insert_sql, insert_params)
    row_id = cur.lastrowid

    cache[cache_key] = row_id
    return row_id, True


def eligible_student_counts_by_school(
    conn: sqlite3.Connection, classroom_id: int
) -> dict[int, int]:
    """Count students eligible to the oraux in a classroom, by school_id."""
    query = """
        SELECT a.school_id, COUNT(*)
        FROM admissions a
        JOIN classroom_students cs ON cs.id = a.classroom_student_id
        WHERE LOWER(TRIM(a.written_status)) = 'admissible'
          AND cs.classroom_id = ?
        GROUP BY a.school_id
    """
    cursor = conn.execute(query, (classroom_id,))
    return dict(cursor.fetchall())


def get_classroom_id_by_year(
    conn: sqlite3.Connection, year: int, branch: Branch
) -> int:
    """Look up an existing classroom by year and branch. Does NOT create one.

    Classrooms are created by import_classroom_students.py; this raises a
    clear error if that step hasn't been run yet for this year and branch.
    """
    row = conn.execute(
        "SELECT id FROM classrooms WHERE year = ? AND branch = ?;", (year, branch)
    ).fetchone()
    if not row:
        raise ValueError(
            f"No classroom found for year {year} and branch {branch}. Run "
            "import_classroom_students.py first."
        )
    return row["id"]


def get_or_create_classroom(
    conn: sqlite3.Connection, year: int, branch: Branch
) -> tuple[int, bool]:
    """Get or create a classroom by exact (year, branch) match.

    No loose matching, consistent with the unique constraint.
    Returns (classroom_id, created: bool).
    """
    return _get_or_create(
        conn,
        "classrooms",
        "WHERE year = ? AND branch = ?",
        (year, branch),
        "INSERT INTO classrooms (year, branch) VALUES (?, ?)",
        (year, branch),
    )


def get_or_create_student(
    conn: sqlite3.Connection, first_name: str, last_name: str
) -> tuple[int, bool]:
    """Get or create a student by exact (first_name, last_name) match.

    No loose matching, consistent with the unique constraint.
    Returns (student_id, created: bool).
    """
    return _get_or_create(
        conn,
        "students",
        "WHERE first_name = ? AND last_name = ?",
        (first_name, last_name),
        "INSERT INTO students (first_name, last_name) VALUES (?, ?)",
        (first_name, last_name),
    )


def get_or_create_classroom_student(
    conn: sqlite3.Connection,
    classroom_id: int,
    student_id: int,
    repeating: bool = False,
) -> tuple[int, bool]:
    """Get or create a classroom student by exact (classroom_id, student_id).

    Matches the unique constraint. Returns (classroom_student_id, created: bool).
    """
    return _get_or_create(
        conn,
        "classroom_students",
        "WHERE classroom_id = ? AND student_id = ?",
        (classroom_id, student_id),
        "INSERT INTO classroom_students (classroom_id, student_id, repeating) VALUES "
        "(?, ?, ?)",
        (classroom_id, student_id, repeating),
    )


def find_bank_by_name(
    conn: sqlite3.Connection, bank_name: str, cache: Cache
) -> int | None:
    """Find a bank by name.

    Loose match (accent/punctuation/case-insensitive) on bank name.
    Returns bank_id, or None if no matching bank exists. Does NOT create.
    """

    cache_key = normalize_name(bank_name)
    return _find_id_loose(conn, "banks", bank_name, cache, cache_key)


def get_or_create_bank(
    conn: sqlite3.Connection, bank_name: str, cache: Cache
) -> tuple[int, bool]:
    """Get or create a bank by name.

    Loose match (accent/punctuation/case-insensitive) on bank name.
    Returns (bank_id, created: bool).
    """
    cache_key = normalize_name(bank_name)

    return _get_or_create_loose(
        conn=conn,
        table="banks",
        name=bank_name,
        cache=cache,
        cache_key=cache_key,
        insert_sql="INSERT INTO banks (name) VALUES (?)",
        insert_params=(bank_name,),
    )


def find_school_by_name(
    conn: sqlite3.Connection, school_name: str, cache: Cache
) -> int | None:
    """Find a school by name, across all banks.

    Loose match (accent/punctuation/case-insensitive) on school name,
    searched across ALL schools regardless of bank (used by imports that
    don't have a bank context, e.g. per-school results files). Returns
    school_id, or None if no matching school exists. Does NOT create.
    School names are globally unique, so there is at most one match.
    """
    cache_key = normalize_name(school_name)
    return _find_id_loose(conn, "schools", school_name, cache, cache_key)


def get_or_create_school(
    conn: sqlite3.Connection, school_name: str, bank_id: int, cache: Cache
) -> tuple[int, bool]:
    """Get or create a school by name within a bank.

    Loose match (accent/punctuation/case-insensitive) on school name,
    scoped to bank_id. Returns (school_id, created: bool).
    """
    cache_key = (bank_id, normalize_name(school_name))

    return _get_or_create_loose(
        conn=conn,
        table="schools",
        name=school_name,
        cache=cache,
        cache_key=cache_key,
        insert_sql="INSERT INTO schools (name, bank_id) VALUES (?, ?)",
        insert_params=(school_name, bank_id),
        where_sql="WHERE bank_id = ?",
        params=(bank_id,),
    )


def upsert_admission_written_result(
    conn: sqlite3.Connection,
    classroom_student_id: int,
    school_id: int,
    written_status: str | None,
    written_points: float | None,
    written_average: float | None,
) -> tuple[int, bool]:
    """Find or create an admissions row by (classroom_student_id, school_id).

    Unlike get_or_create_admission, this OVERWRITES written_status, written_points,
    and written_average if the row already exists — intended for the
    per-school results import, which is the authoritative source for these
    fields (an earlier admissibility import may have already created the
    row with only `written_status` set from a different, earlier-stage value).
    Returns (admission_id, created: bool).
    """
    row = conn.execute(
        "SELECT id FROM admissions WHERE classroom_student_id = ? AND school_id = ?;",
        (classroom_student_id, school_id),
    ).fetchone()

    if row:
        conn.execute(
            """
            UPDATE admissions
            SET written_status = ?, written_points = ?, written_average = ?
            WHERE id = ?;
            """,
            (written_status, written_points, written_average, row[0]),
        )
        return row[0], False

    cur = conn.execute(
        """
        INSERT INTO admissions (
            classroom_student_id, school_id, written_status, written_points,
            written_average
        )
        VALUES (?, ?, ?, ?, ?);
        """,
        (
            classroom_student_id,
            school_id,
            written_status,
            written_points,
            written_average,
        ),
    )
    return cur.lastrowid, True


def find_student_by_name(
    conn: sqlite3.Connection, first_name: str, last_name: str
) -> int | None:
    """Find a student by exact (first_name, last_name).

    Returns student_id, or None if no matching student exists.
    Does NOT create.
    """

    return _find_id(
        conn,
        "students",
        "WHERE first_name = ? AND last_name = ?",
        (first_name, last_name),
    )


def find_classroom_student(
    conn: sqlite3.Connection, classroom_id: int, student_id: int
) -> int | None:
    """Find a classroom student by exact (classroom_id, student_id).

    Returns classroom_student_id, or None if no matching row exists.
    Does NOT create.
    """

    return _find_id(
        conn,
        "classroom_students",
        "WHERE classroom_id = ? AND student_id = ?",
        (classroom_id, student_id),
    )


def find_admission(
    conn: sqlite3.Connection, classroom_student_id: int, school_id: int
) -> sqlite3.Row | None:
    """Find an admission by exact (classroom_student_id, school_id).

    Returns (admission_id, written_points), or None if no matching row
    exists. Does NOT create.
    """
    return conn.execute(
        """
        SELECT id, written_points FROM admissions
        WHERE classroom_student_id = ? AND school_id = ?;
        """,
        (classroom_student_id, school_id),
    ).fetchone()


def update_admission(
    conn: sqlite3.Connection,
    admission_id: int,
    *,
    oral_status: str | None = None,
    rank: int | None = None,
    total_points: float | None = None,
    average: float | None = None,
    oral_points: float | None = None,
) -> None:
    """Updates only the provided attributes of an existing admissions row."""
    fields = {
        "oral_status": oral_status,
        "rank": rank,
        "total_points": total_points,
        "average": average,
        "oral_points": oral_points,
    }
    fields = {k: v for k, v in fields.items() if v is not None}

    if not fields:
        return

    set_clause = ", ".join(f"{col} = ?" for col in fields)
    conn.execute(
        f"UPDATE admissions SET {set_clause} WHERE id = ?;",
        [*fields.values(), admission_id],
    )


def get_school_bank_id(conn: sqlite3.Connection, school_id: int) -> int:
    row = conn.execute(
        "SELECT bank_id FROM schools WHERE id = ?;", (school_id,)
    ).fetchone()
    return row[0]


def get_or_create_admission(
    conn: sqlite3.Connection,
    classroom_student_id: int,
    school_id: int,
    *,
    written_status: str | None = None,
    oral_status: str | None = None,
) -> tuple[int, bool]:
    """Get or create an admission by exact (classroom_student_id, school_id).

    Matches the unique constraint. Returns (admission_id, created: bool).
    Creates only the provided status column(s); does not update status if the
    row already exists.
    """
    fields = {
        "written_status": written_status,
        "oral_status": oral_status,
    }
    fields = {k: v for k, v in fields.items() if v is not None}

    if not fields:
        raise ValueError("at least one of written_status or oral_status is required")

    columns = ["classroom_student_id", "school_id", *fields.keys()]
    placeholders = ", ".join("?" for _ in columns)
    insert_params = (classroom_student_id, school_id, *fields.values())

    return _get_or_create(
        conn=conn,
        table="admissions",
        where_sql="WHERE classroom_student_id = ? AND school_id = ?",
        params=(classroom_student_id, school_id),
        insert_sql=(
            f"INSERT INTO admissions ({', '.join(columns)}) VALUES ({placeholders})"
        ),
        insert_params=insert_params,
    )


def get_or_create_exam(
    conn: sqlite3.Connection,
    exam_name: str,
    bank_id: int,
    format_: ExamFormat,
    cache: Cache,
) -> tuple[int, bool]:
    """Get or create an exam by name within a bank and format.

    Loose match (accent/punctuation/case-insensitive) on exam name,
    scoped to (bank_id, format). Returns (exam_id, created: bool).
    """
    cache_key = (bank_id, format_, normalize_name(exam_name))

    return _get_or_create_loose(
        conn=conn,
        table="exams",
        name=exam_name,
        cache=cache,
        cache_key=cache_key,
        insert_sql="INSERT INTO exams (name, bank_id, format) VALUES (?, ?, ?)",
        insert_params=(exam_name, bank_id, format_),
        where_sql="WHERE bank_id = ? AND format = ?",
        params=(bank_id, format_),
    )


def get_or_create_exam_result(
    conn: sqlite3.Connection,
    classroom_student_id: int,
    exam_id: int,
    points: float | None,
) -> tuple[int, bool]:
    """Get or create an exam result by exact (classroom_student_id, exam_id).

    Matches the unique constraint. Returns (exam_result_id, created: bool).
    Does not update points if it already exists.
    """

    return _get_or_create(
        conn=conn,
        table="exam_results",
        where_sql="WHERE classroom_student_id = ? AND exam_id = ?",
        params=(classroom_student_id, exam_id),
        insert_sql=(
            "INSERT INTO exam_results (classroom_student_id, exam_id, points) "
            "VALUES (?, ?, ?)"
        ),
        insert_params=(classroom_student_id, exam_id, points),
    )
