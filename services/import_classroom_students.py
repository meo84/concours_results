#!/usr/bin/env python3
"""
Import students from ./input/students.xlsx into concours_results.db.

Expected file format:
    - Exactly one sheet
    - Exactly 3 columns with headers: Nom, Prénom, Redoublant
    - One row per student. Redoublant == "R" means repeating, anything else
      (including blank) means not repeating.

Behavior:
    - Missing file -> raise FileNotFoundError naming the expected path.
    - Invalid format -> raise ValueError listing every format issue found.
    - Valid file ->
        - create a classroom for the given year (branch "PC") if none exists yet
        - create a student per row (first_name, last_name) if one doesn't
          already exist for that (first_name, last_name) combination
        - create a classroom_student (classroom_id, repeating, student_id)
          if one doesn't already exist for that (classroom_id, student_id)
          combination

Usage:
    python -m services.import_classroom_students 2026
"""

from pathlib import Path

import argparse
import openpyxl
from services import db_utils
from config import STUDENTS_PATH


EXPECTED_HEADERS = ["Nom", "Prénom", "Redoublant"]
BRANCH = "PC"


def load_and_validate(path: Path):
    """Return rows where rows is a list of (last_name, first_name, repeating).

    Raises FileNotFoundError if the file is missing, ValueError if the
    format is invalid (with every issue found listed in the message).
    """
    if not path.exists():
        raise FileNotFoundError(
            f"Missing input file: {path}. Please place the students file there."
        )

    wb = openpyxl.load_workbook(path, data_only=True)

    errors = []

    if len(wb.sheetnames) != 1:
        errors.append(
            f"Expected a single sheet, found {len(wb.sheetnames)}: {wb.sheetnames}"
        )

    sheet_name = wb.sheetnames[0] if wb.sheetnames else None

    ws = wb[sheet_name] if sheet_name is not None else None
    header_row = None
    if ws is not None:
        header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        # Trim trailing empty columns for a fair comparison
        while header_row and header_row[-1] is None:
            header_row.pop()

        normalized = [str(h).strip() if h is not None else None for h in header_row]
        if normalized != EXPECTED_HEADERS:
            errors.append(
                f"Expected headers {EXPECTED_HEADERS}, found {normalized}."
            )

    if errors:
        raise ValueError("Invalid input file format:\n- " + "\n- ".join(errors))

    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        last_name, first_name, redoublant = (c.value for c in row[:3])
        if last_name is None and first_name is None:
            continue  # skip fully blank rows
        if last_name is None or first_name is None:
            raise ValueError(
                f"Row {row_idx}: missing Nom or Prénom (Nom={last_name!r}, "
                f"Prénom={first_name!r})."
            )
        repeating = str(redoublant).strip().upper() == "R" if redoublant is not None else False
        rows.append((str(last_name).strip(), str(first_name).strip(), repeating))

    return rows


def import_students(conn, classroom_id: int, rows: list) -> tuple:
    """Returns (students_created, students_skipped, classroom_students_created,
    classroom_students_skipped)."""
    students_created = 0
    students_skipped = 0
    classroom_students_created = 0
    classroom_students_skipped = 0

    for last_name, first_name, repeating in rows:
        student_id, student_was_created = db_utils.get_or_create_student(
            conn, first_name, last_name
        )
        if student_was_created:
            students_created += 1
        else:
            students_skipped += 1

        _, cs_was_created = db_utils.get_or_create_classroom_student(
            conn, classroom_id, student_id, repeating
        )
        if cs_was_created:
            classroom_students_created += 1
        else:
            classroom_students_skipped += 1

    return (
        students_created,
        students_skipped,
        classroom_students_created,
        classroom_students_skipped,
    )


def import_classroom_students(year: int) -> None:
    rows = load_and_validate(STUDENTS_PATH)
    print(f"Validated {STUDENTS_PATH}: year={year}, {len(rows)} student row(s).")

    conn = db_utils.get_connection()

    try:
        with conn:
            classroom_id, _created = db_utils.get_or_create_classroom(conn, year, BRANCH)
            print(f"Classroom for year {year} ({BRANCH}): id={classroom_id}")

            created, skipped, cs_created, cs_skipped = import_students(conn, classroom_id, rows)
            print(f"Students: {created} created, {skipped} already existed (skipped).")
            print(
                f"Classroom students: {cs_created} created, {cs_skipped} already existed (skipped)."
            )
    finally:
        conn.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("year", type=int)

    args = parser.parse_args()
    import_classroom_students(args.year)
