#!/usr/bin/env python3
"""
Import written admission statuses from ./input/ecrits/statuts_par_concours/
into concours_results.db.

Expected folder content:
    - At least one .xlsx file.
    - Each filename follows the pattern:
        Statuts pour l_admissibilité de la classe PC-PC pour la banque
        Banque {bank_name} PC{optional more characters}.xlsx
      e.g. "Statuts pour l_admissibilité de la classe PC-PC pour la banque
      Banque CCINP _ Banque e3a - Polytech et Groupe INSA PC(2).xlsx"
      -> bank_name = "CCINP _ Banque e3a - Polytech et Groupe INSA"
    - In each file, the first 3 column headers must be: Numéro, Nom, Prénom
    - Every column header after Prénom is a school name.
      column.

Requires ./input/students.xlsx to have already been imported (via
import_classroom_students.py)

Behavior:
    - Missing or empty folder -> raise FileNotFoundError.
    - Any filename or header format issue -> raise ValueError listing every
      issue found, across all files, before creating anything.
    - No classroom found for the given year -> raise ValueError
      telling you to run import_classroom_students.py first.
    - Valid folder ->
        - create a bank per distinct bank_name (loose match: same words,
          case-insensitive, accent- and punctuation-insensitive count as
          the same bank)
        - create a school per school-name column, scoped to its bank
          (same loose matching), only if it doesn't already exist
        - for each data row, find or create the student (exact match on
          first_name/last_name), find or create the classroom_student
          (exact match on classroom_id and student_id)
        - create the admissions record (classroom_student_id,
          school_id), setting status from the school-name column if the status
          is not blank (no admissions record created otherwise)

Usage:
    python -m services.import_written_admission_statuses 2026
"""

import re
import unicodedata
from pathlib import Path

import openpyxl

from services import db_utils, excel_utils
from config import WRITTEN_STATUSES_PER_BANK_PATH

EXPECTED_FIRST_HEADERS = ["Numéro", "Nom", "Prénom"]

FILENAME_PATTERN = re.compile(
    r"^Statuts\s+pour\s+l_admissibilité\s+de\s+la\s+classe\s+PC-PC\s+pour\s+la\s+banque\s+"
    r"Banque\s+(?P<bank_name>.+)\s+PC(?P<suffix>.*)\.xlsx$"
)

def validate_and_read_file(path: Path):
    """Return (bank_name, school_names, data_rows, errors).

    data_rows is a list of (row_idx, last_name, first_name, status) where
    status is a list aligned with school_names (None for a blank cell).
    On any format error, returns (bank_name_or_None, None, None, errors).
    """
    bank_name, error = excel_utils.validate_filename(
        path, FILENAME_PATTERN, "bank_name",
        "Status pour l_admissibilité de la classe PC-PC pour la banque Banque {bank_name} PC...xlsx",
    )

    if error:
        return None, None, None, [error]

    all_rows = excel_utils.read_xlsx_rows(path)

    empty_file_error = excel_utils.validate_non_empty_file(all_rows, path.name)
    if empty_file_error:
        return bank_name, None, None, [empty_file_error]

    normalized_headers = excel_utils.normalize_header_row(all_rows[0])

    if len(normalized_headers) < 3 or normalized_headers[:3] != EXPECTED_FIRST_HEADERS:
        error = f"{path.name}: expected first 3 headers {EXPECTED_FIRST_HEADERS}, "
        f"found {normalized_headers[:3]}."
        return bank_name, None, None, [error]

    school_names = [h for h in normalized_headers[3:] if h]
    if not school_names:
        error = f"{path.name}: no school columns found after 'Prénom'."
        return bank_name, None, None, [error]

    errors = []
    n_schools = len(school_names)
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
        status = list(row[3:3 + n_schools]) + [None] * max(0, n_schools - len(row[3:]))
        data_rows.append((row_idx, str(last_name).strip(), str(first_name).strip(), status))

    if errors:
        return bank_name, school_names, None, errors

    return bank_name, school_names, data_rows, errors


def import_written_admission_statuses(year: int) -> None:
    files = excel_utils.discover_files(WRITTEN_STATUSES_PER_BANK_PATH)

    all_errors = []
    parsed = []  # list of (path, bank_name, school_names, data_rows)
    for path in files:
        bank_name, school_names, data_rows, errors = validate_and_read_file(path)
        if errors:
            all_errors.extend(errors)
        else:
            parsed.append((path, bank_name, school_names, data_rows))

    if all_errors:
        raise ValueError("Invalid input format:\n- " + "\n- ".join(all_errors))

    conn = db_utils.get_connection()

    try:
        with conn:
            classroom_id = db_utils.get_classroom_id_by_year(conn, year)
            print(f"Using classroom for year {year} (id={classroom_id})")

            bank_cache = {}
            school_cache = {}
            banks_created = 0
            schools_created = 0
            schools_skipped = 0
            students_created = 0
            students_skipped = 0
            classroom_students_created = 0
            classroom_students_skipped = 0
            admissions_created = 0
            admissions_skipped = 0

            for path, bank_name, school_names, data_rows in parsed:
                bank_id, created = db_utils.get_or_create_bank(conn, bank_name, bank_cache)
                if created:
                    banks_created += 1
                    print(f"Created bank: {bank_name!r} (id={bank_id})")

                school_ids = []
                for school_name in school_names:
                    school_id, created = db_utils.get_or_create_school(conn, school_name, bank_id, school_cache)
                    school_ids.append(school_id)
                    if created:
                        schools_created += 1
                        print(f"  Created school: {school_name!r} (id={school_id}, bank_id={bank_id})")
                    else:
                        schools_skipped += 1

                for row_idx, last_name, first_name, statuses_row in data_rows:
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

                    for school_id, status in zip(school_ids, statuses_row):
                        if status is None:
                            continue  # no status for this student/school
                        _, created = db_utils.get_or_create_admission(
                            conn, classroom_student_id, school_id, status
                        )
                        if created:
                            admissions_created += 1
                        else:
                            admissions_skipped += 1

    finally:
        conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed.\n"
        f"Banks created: {banks_created}.\n"
        f"Schools created: {schools_created}, skipped: {schools_skipped}.\n"
        f"Students created: {students_created}, skipped: {students_skipped}.\n"
        f"Classroom students created: {classroom_students_created}, skipped: {classroom_students_skipped}.\n"
        f"Admissions created: {admissions_created}, skipped: {admissions_skipped}.\n"
    )


if __name__ == "__main__":
    import_written_admission_statuses()
