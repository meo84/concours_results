"""
Shared helpers for the concours_results import scripts.

Keeping these here means each import_*.py script stays a thin, independently
runnable step (matching how you run them in sequence), while sharing the
same DB-access and get-or-create logic instead of duplicating it.
"""

from config import DB_PATH
import sqlite3
import unicodedata


def get_connection() -> sqlite3.Connection:
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found at {DB_PATH}. Run init_db.py first.")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def normalize_name(name: str) -> str:
    """Lowercase, strip accents and punctuation, collapse whitespace for loose matching."""
    decomposed = unicodedata.normalize("NFKD", name.strip().lower())
    without_accents = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    cleaned = "".join(ch if ch.isalnum() or ch.isspace() else " " for ch in without_accents)
    return " ".join(cleaned.split())


def _exact_match_get_or_create(
    conn: sqlite3.Connection,
    table: str,
    scope_sql: str,
    scope_params: tuple = (),
    insert_sql: str | None = None,
    insert_params: tuple = (),
) -> tuple[int | None, bool]:
    """Exact match, optionally create if no match exists.

    Returns:
        (row_id, False) if an existing row is found
        (row_id, True) if a new row is created
        (None, False) if no row is found and insert_sql is not provided
    """
    row = conn.execute(f"SELECT id FROM {table} {scope_sql}", scope_params).fetchone()

    if row:
        return row[0], False

    if insert_sql is None:
        return None, False

    cur = conn.execute(insert_sql, insert_params)
    return cur.lastrowid, True


def _loose_match_get_or_create(
    conn: sqlite3.Connection,
    table: str,
    name: str,
    cache: dict,
    cache_key,
    scope_sql: str = "",
    scope_params: tuple = (),
    insert_sql: str | None = None,
    insert_params: tuple = (),
) -> tuple[int | None, bool]:
    """Loose match (accent/punctuation/case-insensitive) on name, optionally scoped.

    Returns:
        (row_id, False) if an existing row is found
        (row_id, True) if a new row is created
        (None, False) if no row is found and insert_sql is not provided
    """
    if cache_key in cache:
        return cache[cache_key], False

    target = normalize_name(name)

    rows = conn.execute(f"SELECT id, name FROM {table} {scope_sql}", scope_params).fetchall()

    for row_id, existing_name in rows:
        if normalize_name(existing_name) == target:
            cache[cache_key] = row_id
            return row_id, False

    # No match: either create or return None
    if insert_sql is None:
        return None, False

    cur = conn.execute(insert_sql, insert_params)
    row_id = cur.lastrowid

    cache[cache_key] = row_id
    return row_id, True

def get_classroom_id_by_year(conn: sqlite3.Connection, year: int) -> int:
    """Look up an existing classroom by year. Does NOT create one.

    Classrooms are created by import_classroom_students.py; this raises a
    clear error if that step hasn't been run yet for this year.
    """
    row = conn.execute("SELECT id FROM classrooms WHERE year = ?;", (year,)).fetchone()
    if not row:
        raise ValueError(
            f"No classroom found for year {year}. Run import_classroom_students.py first."
        )
    return row["id"]


def get_or_create_classroom(conn: sqlite3.Connection, year: int, branch: str) -> tuple:
    """Exact match on year — no loose matching, consistent
    with the unique constraint. Returns (classroom_id, created: bool)."""
    return  _exact_match_get_or_create(
        conn, "classrooms", "WHERE year = ?", (year,),
        "INSERT INTO classrooms (year, branch) VALUES (?, ?)", (year, branch),
    )


def get_or_create_student(
    conn: sqlite3.Connection, first_name: str, last_name: str
) -> tuple:
    """Exact match on (first_name, last_name) — no loose matching, consistent
    with the unique constraint. Returns (student_id, created: bool)."""
    return _exact_match_get_or_create(
        conn, "students", "WHERE first_name = ? AND last_name = ?", (first_name, last_name),
        "INSERT INTO students (first_name, last_name) VALUES (?, ?)", (first_name, last_name),
    )


def get_or_create_classroom_student(
    conn: sqlite3.Connection,
    classroom_id: int,
    student_id: int,
    repeating: bool = False,
) -> tuple:
    """Exact match on (classroom_id, student_id), matching the unique
    constraint. Returns (classroom_student_id, created: bool)."""
    return _exact_match_get_or_create(
        conn, "classroom_students", "WHERE classroom_id = ? AND student_id = ?", (classroom_id, student_id),
        "INSERT INTO classroom_students (classroom_id, student_id, repeating) VALUES (?, ?, ?)", (classroom_id, student_id, repeating),
    )


def find_bank_by_name(conn: sqlite3.Connection, bank_name: str, cache: dict):
    """Loose match (accent/punctuation/case-insensitive) on bank name.
    Returns bank_id, or None if no matching bank exists. Does NOT create."""

    cache_key = normalize_name(bank_name)
    bank_id, _created = _loose_match_get_or_create(conn=conn, table="banks", name=bank_name, cache=cache, cache_key=cache_key)

    return bank_id


def get_or_create_bank(conn: sqlite3.Connection, bank_name: str, cache: dict) -> tuple:
    """Loose match (accent/punctuation/case-insensitive) on bank name.
    Returns (bank_id, created: bool)."""
    cache_key = normalize_name(bank_name)

    return _loose_match_get_or_create(
        conn=conn, table="banks", name=bank_name, cache=cache, cache_key=cache_key,
        insert_sql="INSERT INTO banks (name) VALUES (?)", insert_params=(bank_name,)
    )


def find_school_by_name(conn: sqlite3.Connection, school_name: str, cache: dict):
    """Loose match (accent/punctuation/case-insensitive) on school name,
    searched across ALL schools regardless of bank (used by imports that
    don't have a bank context, e.g. per-school results files). Returns
    school_id, or None if no matching school exists. Does NOT create.
    If multiple schools across different banks share the same loose name,
    the first match found is returned — this isn't disambiguated further."""
    cache_key = normalize_name(school_name)
    school_id, _created = _loose_match_get_or_create(conn=conn, table="schools", name=school_name, cache=cache, cache_key=cache_key)

    return school_id


def get_or_create_school(
    conn: sqlite3.Connection, school_name: str, bank_id: int, cache: dict
) -> tuple:
    """Loose match (accent/punctuation/case-insensitive) on school name,
    scoped to bank_id. Returns (school_id, created: bool)."""
    cache_key = (bank_id, normalize_name(school_name))

    return _loose_match_get_or_create(
        conn=conn, table="schools", name=school_name, cache=cache, cache_key=cache_key,
        scope_sql="WHERE bank_id = ?", scope_params=(bank_id,),
        insert_sql="INSERT INTO schools (name, bank_id) VALUES (?, ?)", insert_params=(school_name, bank_id),
    )



def get_or_create_admission(
    conn: sqlite3.Connection, classroom_student_id: int, school_id: int, status: str
) -> tuple:
    """Exact match on (classroom_student_id, school_id), matching the unique
    constraint. Returns (admission_id, created: bool). Does not update status
    if it already exists."""

    return _exact_match_get_or_create(
        conn=conn, table="admissions", scope_sql="WHERE classroom_student_id = ? AND school_id = ?", scope_params=(classroom_student_id, school_id),
        insert_sql="INSERT INTO admissions (classroom_student_id, school_id, status) VALUES (?, ?, ?)", insert_params=(classroom_student_id, school_id, status),
    )


def upsert_admission_written_result(
    conn: sqlite3.Connection,
    classroom_student_id: int,
    school_id: int,
    status,
    written_points,
    written_average,
) -> tuple:
    """Find or create an admissions row by (classroom_student_id, school_id).

    Unlike get_or_create_admission, this OVERWRITES status, written_points,
    and written_average if the row already exists — intended for the
    per-school results import, which is the authoritative source for these
    fields (an earlier admissibility import may have already created the
    row with only `status` set from a different, earlier-stage value).
    Returns (admission_id, created: bool)."""
    row = conn.execute(
        "SELECT id FROM admissions WHERE classroom_student_id = ? AND school_id = ?;",
        (classroom_student_id, school_id),
    ).fetchone()

    if row:
        conn.execute(
            """
            UPDATE admissions
            SET status = ?, written_points = ?, written_average = ?
            WHERE id = ?;
            """,
            (status, written_points, written_average, row[0]),
        )
        return row[0], False

    cur = conn.execute(
        """
        INSERT INTO admissions (classroom_student_id, school_id, status, written_points, written_average)
        VALUES (?, ?, ?, ?, ?);
        """,
        (classroom_student_id, school_id, status, written_points, written_average),
    )
    return cur.lastrowid, True


def find_student_by_name(conn: sqlite3.Connection, first_name: str, last_name: str):
    """Exact match on (first_name, last_name). Returns student_id, or None
    if no matching student exists. Does NOT create."""

    student_id, _created =  _exact_match_get_or_create(
        conn=conn, table="students", scope_sql="WHERE first_name = ? AND last_name = ?", scope_params=(first_name, last_name),
    )
    return student_id


def find_classroom_student(conn: sqlite3.Connection, classroom_id: int, student_id: int):
    """Exact match on (classroom_id, student_id). Returns classroom_student_id,
    or None if no matching row exists. Does NOT create."""

    classroom_student_id, _created =  _exact_match_get_or_create(
        conn=conn, table="classroom_students", scope_sql="WHERE classroom_id = ? AND student_id = ?", scope_params=(classroom_id, student_id),
    )
    return classroom_student_id


def find_admission(conn: sqlite3.Connection, classroom_student_id: int, school_id: int):
    """Exact match on (classroom_student_id, school_id). Returns
    (admission_id, written_points), or None if no matching row exists.
    Does NOT create."""
    row = conn.execute(
        """
        SELECT id, written_points FROM admissions
        WHERE classroom_student_id = ? AND school_id = ?;
        """,
        (classroom_student_id, school_id),
    ).fetchone()
    return row if row else None


def update_admission_oral_result(
    conn: sqlite3.Connection,
    admission_id: int,
    status,
    rank,
    total_points,
    average,
    oral_points,
) -> None:
    """Overwrites status, rank, total_points, average, and oral_points on an
    existing admissions row. Does not touch written_points/written_average."""
    conn.execute(
        """
        UPDATE admissions
        SET status = ?, rank = ?, total_points = ?, average = ?, oral_points = ?
        WHERE id = ?;
        """,
        (status, rank, total_points, average, oral_points, admission_id),
    )


def get_school_bank_id(conn: sqlite3.Connection, school_id: int) -> int:
    row = conn.execute("SELECT bank_id FROM schools WHERE id = ?;", (school_id,)).fetchone()
    return row[0]


def get_or_create_admission(
    conn: sqlite3.Connection, classroom_student_id: int, school_id: int, status: str
) -> tuple:
    """Exact match on (classroom_student_id, school_id), matching the unique
    constraint. Returns (admission_id, created: bool). Does not update
    status if it already exists."""

    return _exact_match_get_or_create(
        conn=conn, table="admissions", scope_sql="WHERE classroom_student_id = ? AND school_id = ?", scope_params=(classroom_student_id, school_id),
        insert_sql="INSERT INTO admissions (classroom_student_id, school_id, status) VALUES (?, ?, ?)", insert_params=(classroom_student_id, school_id, status),
    )


def get_or_create_exam(
    conn: sqlite3.Connection, exam_name: str, bank_id: int, format_: str, cache: dict
) -> tuple:
    """Loose match (accent/punctuation/case-insensitive) on exam name,
    scoped to (bank_id, format). Returns (exam_id, created: bool)."""
    cache_key = (bank_id, format_, normalize_name(exam_name))

    return _loose_match_get_or_create(
        conn=conn, table="exams", name=exam_name, cache=cache, cache_key=cache_key,
        scope_sql="WHERE bank_id = ? AND format = ?", scope_params=(bank_id, format_),
        insert_sql="INSERT INTO exams (name, bank_id, format) VALUES (?, ?, ?)", insert_params=(exam_name, bank_id, format_),
    )


def get_or_create_exam_result(
    conn: sqlite3.Connection, classroom_student_id: int, exam_id: int, points
) -> tuple:
    """Exact match on (classroom_student_id, exam_id), matching the unique
    constraint. Returns (exam_result_id, created: bool). Does not update
    points if it already exists."""

    return _exact_match_get_or_create(
        conn=conn, table="exam_results", scope_sql="WHERE classroom_student_id = ? AND exam_id = ?", scope_params=(classroom_student_id, exam_id),
        insert_sql="INSERT INTO exam_results (classroom_student_id, exam_id, points) VALUES (?, ?, ?)", insert_params=(classroom_student_id, exam_id, points),
    )
