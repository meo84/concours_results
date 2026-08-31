#!/usr/bin/env python3
"""
Import exams and exam results from ./input/ecrits/notes_par_concours/ into
concours_results.db.

Expected folder content:
    - At least one .xlsx file.
    - Each filename follows the pattern:
        Résultats de la classe PC-PC pour la Banque {bank_name} PC{optional more characters}.xlsx
      e.g. "Résultats de la classe PC-PC pour la Banque Mines Ponts PC(1).xlsx"
      -> bank_name = "Mines Ponts"
    - In each file, the first 3 column headers must be: Numéro, Nom, Prénom
    - Every column header after Prénom is an exam name.
    - Each data row holds one student's points per exam column.

Requires ./input/students.xlsx to have already been imported (via
import_classroom_students.py)

Behavior:
    - Missing or empty folder -> raise FileNotFoundError.
    - Any filename or header format issue -> raise ValueError listing every
      issue found, across all files, before creating anything.
    - No classroom found for the given year -> raise ValueError
      telling you to run import_classroom_students.py first.
    - Valid folder ->
        - for each file, find the bank by name (loose match: same words,
          case-insensitive, accent- and punctuation-insensitive count as
          the same bank). If no matching bank exists, the file is skipped
          and a warning is recorded (printed in the summary at the end) —
          processing continues with the remaining files.
        - for each column header after Prénom, create an exam (name,
          bank_id, format="written") if one doesn't already exist for that
          (bank_id, name, format) combination (same loose name matching)
        - for each data row, find or create the student (exact match on
          first_name/last_name), find or create the classroom_student
          (exact match on classroom_id and student_id)
        - for each exam column with a non-blank cell, find or create an
          exam_results record (classroom_student_id, exam_id, points).
          Blank cells are skipped (no exam_results record created).

Usage:
    python -m services.import_written_exam_results 2026
"""

import argparse
import re
import unicodedata
from pathlib import Path

import openpyxl

from services import db_utils, excel_utils
from config import WRITTEN_EXAM_RESULTS_PATH

EXPECTED_FIRST_HEADERS = ["Numéro", "Nom", "Prénom"]
FORMAT = "written"

FILENAME_PATTERN = re.compile(
    r"^Résultats\s+de\s+la\s+classe\s+PC-PC\s+pour\s+la\s+Banque\s+"
    r"(?P<bank_name>.+)\s+PC(?P<suffix>.*)\.xlsx$"
)


def validate_and_read_file(path: Path):
    """Return (bank_name, exam_names, data_rows, errors).

    data_rows is a list of (row_idx, last_name, first_name, points) where
    points is a list aligned with exam_names (None for a blank cell).
    On any format error, returns (bank_name_or_None, None, None, errors).
    """
    errors = []

    bank_name, error = excel_utils.validate_filename(
        path, FILENAME_PATTERN, "bank_name",
        "Résultats de la classe PC-PC pour la Banque {{bank_name}} PC...xlsx",
    )

    if error:
        return None, None, None, [error]

    all_rows = excel_utils.read_xlsx_rows(path)

    empty_file_error = excel_utils.validate_non_empty_file(all_rows, path.name)
    if empty_file_error:
        return bank_name, None, None, [empty_file_error]

    normalized_headers = excel_utils.normalize_header_row(all_rows[0])

    if len(normalized_headers) < 3 or normalized_headers[:3] != EXPECTED_FIRST_HEADERS:
        errors.append(
            f"{path.name}: expected first 3 headers {EXPECTED_FIRST_HEADERS}, "
            f"found {normalized_headers[:3]}."
        )
        return bank_name, None, None, errors

    exam_names = [h for h in normalized_headers[3:] if h]
    if not exam_names:
        errors.append(f"{path.name}: no exam columns found after 'Prénom'.")
        return bank_name, None, None, errors

    n_exams = len(exam_names)
    data_rows = []
    for row_idx, row in enumerate(all_rows[1:], start=2):
        last_name = row[1] if len(row) > 1 else None
        first_name = row[2] if len(row) > 2 else None
        if last_name is None and first_name is None:
            continue  # skip fully blank rows
        name_columns_error = excel_utils.validate_name_columns(row_idx, last_name, first_name, path.name)
        if name_columns_error:
            errors.append(name_columns_error)
            continue
        points = list(row[3:3 + n_exams]) + [None] * max(0, n_exams - len(row[3:]))
        data_rows.append((row_idx, str(last_name).strip(), str(first_name).strip(), points))

    if errors:
        return bank_name, exam_names, None, errors

    return bank_name, exam_names, data_rows, errors


def import_written_exam_results(year: int) -> None:
    files = excel_utils.discover_files(WRITTEN_EXAM_RESULTS_PATH)

    all_errors = []
    parsed = []  # list of (path, bank_name, exam_names, data_rows)
    for path in files:
        bank_name, exam_names, data_rows, errors = validate_and_read_file(path)
        if errors:
            all_errors.extend(errors)
        else:
            parsed.append((path, bank_name, exam_names, data_rows))

    if all_errors:
        raise ValueError("Invalid input format:\n- " + "\n- ".join(all_errors))

    conn = db_utils.get_connection()

    try:
        with conn:
            classroom_id = db_utils.get_classroom_id_by_year(conn, year)
            print(f"Using classroom for year {year} (id={classroom_id})")

            bank_cache = {}
            exam_cache = {}
            bank_not_found_warnings = []
            exams_created = 0
            exams_skipped = 0
            students_created = 0
            students_skipped = 0
            classroom_students_created = 0
            classroom_students_skipped = 0
            results_created = 0
            results_skipped = 0

            for path, bank_name, exam_names, data_rows in parsed:
                bank_id = db_utils.find_bank_by_name(conn, bank_name, bank_cache)
                if bank_id is None:
                    bank_not_found_warnings.append(
                        f"{path.name}: no bank found matching {bank_name!r}. File skipped."
                    )
                    continue

                exam_ids = []
                for exam_name in exam_names:
                    exam_id, created = db_utils.get_or_create_exam(conn, exam_name, bank_id, FORMAT, exam_cache)
                    exam_ids.append(exam_id)
                    if created:
                        exams_created += 1
                        print(f"Created exam: {exam_name!r} (id={exam_id}, bank_id={bank_id})")
                    else:
                        exams_skipped += 1

                for row_idx, last_name, first_name, points_row in data_rows:
                    student_id, created = db_utils.get_or_create_student(conn, first_name, last_name)
                    if created:
                        students_created += 1
                    else:
                        students_skipped += 1

                    classroom_student_id, created = db_utils.get_or_create_classroom_student(
                        conn, classroom_id, student_id
                    )
                    if created:
                        classroom_students_created += 1
                    else:
                        classroom_students_skipped += 1

                    for exam_id, points in zip(exam_ids, points_row):
                        if points is None:
                            continue  # no result for this student/exam
                        _, created = db_utils.get_or_create_exam_result(
                            conn, classroom_student_id, exam_id, points
                        )
                        if created:
                            results_created += 1
                        else:
                            results_skipped += 1

    finally:
        conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed.\n"
        f"Exams created: {exams_created}, skipped: {exams_skipped}.\n"
        f"Students created: {students_created}, skipped: {students_skipped}.\n"
        f"Classroom students created: {classroom_students_created}, skipped: {classroom_students_skipped}.\n"
        f"Exam results created: {results_created}, skipped: {results_skipped}."
    )

    if bank_not_found_warnings:
        print("\nWarnings:")
        for w in bank_not_found_warnings:
            print(f"- {w}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("year", type=int)

    args = parser.parse_args()
    import_written_exam_results(args.year)
