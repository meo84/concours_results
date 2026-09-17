#!/usr/bin/env python3
"""
Import oral admission statuses from ./input/oraux/statuts_par_concours/
into concours_results.db.

Expected folder content:
    - At least one .xlsx file.
    - Each filename follows the pattern:
        Statuts pour l_admission de la classe PC-PC pour la banque
        Banque {bank_name} PC{optional more characters}.xlsx
      e.g. "Statuts pour l_admission de la classe PC-PC pour la banque
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
        - for each file, find the bank by name (loose match: same words,
          case-insensitive, accent- and punctuation-insensitive count as
          the same school). If no matching bank exists, the file is
          skipped and a warning is recorded (printed in the summary at the
          end) — processing continues with the remaining files.
        - for each data row:
            - find the student (exact match on first_name/last_name) and
              the classroom_student (exact match on classroom_id,
              student_id). Neither is created here — both are expected to
              already exist from earlier imports. If either is missing,
              the row is skipped and an error is recorded (printed at the
              end) — processing continues with the remaining rows.
            - find the admissions record (classroom_student_id, school_id).
              If missing, the admissions update for this row is skipped and
              an error is recorded.
              If found, it is updated (overwritten, not just filled in) with
              the oral_ status from the school-name column if the status is not blank

Usage:
    python -m services.import_oral_admission_statuses 2026
"""

import re
import unicodedata
from pathlib import Path

import openpyxl

from services import db_utils, excel_utils
from config import ORAL_STATUSES_PER_BANK_PATH

EXPECTED_FIRST_HEADERS = ["Numéro", "Nom", "Prénom"]

FILENAME_PATTERN = re.compile(
    r"^Statuts\s+pour\s+l_admission\s+de\s+la\s+classe\s+PC-PC\s+pour\s+la\s+banque\s+"
    r"Banque\s+(?P<bank_name>.+)\s+PC(?P<suffix>.*)\.xlsx$"
)

def validate_and_read_file(path: Path):
    """Return (bank_name, school_names, data_rows, errors).

    data_rows is a list of (row_idx, last_name, first_name, oral_status) where
    oral_status is a list aligned with school_names (None for a blank cell).
    On any format error, returns (bank_name_or_None, None, None, errors).
    """
    bank_name, error = excel_utils.validate_filename(
        path, FILENAME_PATTERN, "bank_name",
        "Status pour l_admission de la classe PC-PC pour la banque Banque {bank_name} PC...xlsx",
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
        oral_status = list(row[3:3 + n_schools]) + [None] * max(0, n_schools - len(row[3:]))
        data_rows.append((row_idx, str(last_name).strip(), str(first_name).strip(), oral_status))

    if errors:
        return bank_name, school_names, None, errors

    return bank_name, school_names, data_rows, errors


def import_oral_admission_statuses(year: int) -> None:
    files = excel_utils.discover_files(ORAL_STATUSES_PER_BANK_PATH)

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
            school_not_found_warnings = []
            bank_not_found_warnings = []
            row_errors = []
            admissions_updated = 0

            for path, bank_name, school_names, data_rows in parsed:
                bank_id = db_utils.find_bank_by_name(conn, bank_name, bank_cache)
                if bank_id is None:
                    bank_not_found_warnings.append(
                        f"{path.name}: no bank found matching {bank_name!r}. File skipped."
                    )
                    continue

                school_ids = []
                for school_name in school_names:
                    school_id = db_utils.find_school_by_name(conn, school_name, school_cache)
                    if school_id is None:
                        school_not_found_warnings.append(
                            f"{path.name}: no school found matching {school_name!r}. File skipped."
                        )
                        continue

                    school_ids.append(school_id)

                for row_idx, last_name, first_name, oral_statuses_row in data_rows:
                    student_id = db_utils.find_student_by_name(conn, first_name, last_name)
                    if student_id is None:
                        row_errors.append(
                            f"{path.name} row {row_idx}: no student found for "
                            f"{first_name!r} {last_name!r}. Row skipped."
                        )
                        continue

                    classroom_student_id = db_utils.find_classroom_student(conn, classroom_id, student_id)
                    if classroom_student_id is None:
                        row_errors.append(
                            f"{path.name} row {row_idx}: no classroom_student found for "
                            f"{first_name!r} {last_name!r} in classroom {classroom_id}. Row skipped."
                        )
                        continue

                    for school_id, oral_status in zip(school_ids, oral_statuses_row):
                        if oral_status is None:
                            continue  # no oral_status for this student/school

                        admission = db_utils.find_admission(conn, classroom_student_id, school_id)
                        if admission is None:
                            row_errors.append(
                                f"{path.name} row {row_idx}: no admissions record found for "
                                f"{first_name!r} {last_name!r} at school_id={school_id}. "
                                f"Admissions update skipped for this row."
                            )
                            continue

                        admission_id, _ = admission
                        oral_status, rank = parse_admission_status(oral_status)
                        db_utils.update_admission(conn, admission_id, oral_status=oral_status, rank=rank)
                        admissions_updated += 1

    finally:
        conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed\n"
        f"Admissions updated: {admissions_updated}.\n"
    )

def parse_admission_status(status):
    normalized = db_utils.normalize_name(status)
    match = re.match(r"^classe e (\d+)$", normalized)
    if match:
        return "classé-e", int(match.group(1))
    return status, None


if __name__ == "__main__":
    import_oral_admission_statuses()
