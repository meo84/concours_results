#!/usr/bin/env python3
"""
Import per-school admission results from ./input/ecrits/notes_par_ecole/
into concours_results.db.

Expected folder content:
    - At least one .xlsx file.
    - Each filename follows the pattern:
        Résultats de l_admissibilité pour le {school_name} PC{optional more characters}.xlsx
      e.g. "Resultats de l_admissibilité pour le Arts et Métiers de PC (2).xlsx"
      -> school_name = "Arts et Métiers"
      (the accent on the first word is tolerated either way: "Résultats" or
      "Resultats", since both appear in practice)
    - In each file, columns must be, in order:
        Numéro, Nom, Prénom, Statut   (strict, exact header match)
        Total écrit, Moyenne          (loose header match: accent/
                                        punctuation/case-insensitive)
      Any columns after these 6 are ignored (extra columns are allowed).
    - Each data row holds one student's written admission result at that
      school.

Requires ./input/students.xlsx to have already been imported (via
import_classroom_students.py)

Behavior:
    - Missing or empty folder -> raise FileNotFoundError.
    - Any filename or header format issue -> raise ValueError listing every
      issue found, across all files, before creating anything.
    - No classroom found for the given year -> raise ValueError
      telling you to run import_classroom_students.py first.
    - Valid folder ->
        - for each file, find the school by name (loose match: same words,
          case-insensitive, accent- and punctuation-insensitive count as
          the same school). If no matching school exists, the file is
          skipped and a warning is recorded (printed in the summary at the
          end) — processing continues with the remaining files.
        - for each data row, find or create the student (exact match on
          first_name/last_name), find or create the classroom_student
          (exact match on classroom_id and student_id)
        - create the admissions record (classroom_student_id,
          school_id), setting status (from Statut), written_points (from
          Total écrit), and written_average (from Moyenne).
          Rows where all three of Statut/Total écrit/Moyenne are blank
          are skipped entirely (no admissions record created).

Usage:
    python import_written_admissions.py 2026
"""

import argparse
import re
import unicodedata
from pathlib import Path

import openpyxl

from services import common
from services import import_classroom_students
from config import WRITTEN_ADMISSIONS_PER_SCHOOL_PATH

EXPECTED_STRICT_HEADERS = ["Numéro", "Nom", "Prénom", "Statut"]
EXPECTED_LOOSE_HEADERS = ["Total écrit", "Moyenne"]

FILENAME_PATTERN = re.compile(
    r"^R[eé]sultats\s+de\s+l_admissibilit[eé]\s+pour\s+le\s+(?P<school_name>.+)\s+de\s+PC(?P<suffix>.*)\.xlsx$"
)


def validate_and_read_file(path: Path):
    """Return (school_name, data_rows, errors).

    data_rows is a list of (row_idx, last_name, first_name, status,
    written_points, written_average). On any format error, returns
    (school_name_or_None, None, errors).
    """
    errors = []

    normalized_filename = unicodedata.normalize("NFC", path.name)
    match = FILENAME_PATTERN.match(normalized_filename)
    if not match:
        errors.append(
            f"{path.name}: filename does not match the expected pattern "
            f"'Résultats de l_admissibilité pour le {{school_name}} de PC...xlsx'. "
            f"(raw: {normalized_filename!r})"
        )
        return None, None, errors

    school_name = match.group("school_name").strip()

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))

    if not all_rows:
        errors.append(f"{path.name}: file is empty.")
        return school_name, None, errors

    header_row = list(all_rows[0])
    while header_row and header_row[-1] is None:
        header_row.pop()
    normalized_headers = [str(h).strip() if h is not None else None for h in header_row]

    expected_count = len(EXPECTED_STRICT_HEADERS) + len(EXPECTED_LOOSE_HEADERS)
    if len(normalized_headers) < expected_count:
        errors.append(
            f"{path.name}: expected at least {expected_count} columns starting with "
            f"{EXPECTED_STRICT_HEADERS + EXPECTED_LOOSE_HEADERS}, found {normalized_headers}."
        )
        return school_name, None, errors

    strict_actual = normalized_headers[: len(EXPECTED_STRICT_HEADERS)]
    if strict_actual != EXPECTED_STRICT_HEADERS:
        errors.append(
            f"{path.name}: expected first headers {EXPECTED_STRICT_HEADERS}, found {strict_actual}."
        )
        return school_name, None, errors

    loose_start = len(EXPECTED_STRICT_HEADERS)
    loose_end = loose_start + len(EXPECTED_LOOSE_HEADERS)
    loose_actual = normalized_headers[loose_start:loose_end]
    for expected, actual in zip(EXPECTED_LOOSE_HEADERS, loose_actual):
        if common.normalize_name(actual or "") != common.normalize_name(expected):
            errors.append(
                f"{path.name}: expected header {expected!r} (loose match), found {actual!r}."
            )
    if errors:
        return school_name, None, errors

    data_rows = []
    for row_idx, row in enumerate(all_rows[1:], start=2):
        last_name = row[1] if len(row) > 1 else None
        first_name = row[2] if len(row) > 2 else None
        status = row[3] if len(row) > 3 else None
        written_points = row[4] if len(row) > 4 else None
        written_average = row[5] if len(row) > 5 else None

        if last_name is None and first_name is None:
            continue  # skip fully blank rows
        if last_name is None or first_name is None:
            errors.append(
                f"{path.name} row {row_idx}: missing Nom or Prénom "
                f"(Nom={last_name!r}, Prénom={first_name!r})."
            )
            continue

        data_rows.append((
            row_idx,
            str(last_name).strip(),
            str(first_name).strip(),
            status,
            written_points,
            written_average,
        ))

    if errors:
        return school_name, None, errors

    return school_name, data_rows, errors


def import_written_admissions(year: int) -> None:
    files = common.discover_files(WRITTEN_ADMISSIONS_PER_SCHOOL_PATH)

    all_errors = []
    parsed = []  # list of (path, school_name, data_rows)
    for path in files:
        school_name, data_rows, errors = validate_and_read_file(path)
        if errors:
            all_errors.extend(errors)
        else:
            parsed.append((path, school_name, data_rows))

    if all_errors:
        raise ValueError("Invalid input format:\n- " + "\n- ".join(all_errors))

    conn = common.get_connection()

    classroom_id = common.get_classroom_id_by_year(conn, year)
    print(f"Using classroom for year {year} (id={classroom_id})")

    school_cache = {}
    school_not_found_warnings = []
    students_created = 0
    students_skipped = 0
    classroom_students_created = 0
    classroom_students_skipped = 0
    admissions_created = 0
    admissions_updated = 0
    rows_skipped_blank = 0

    for path, school_name, data_rows in parsed:
        school_id = common.find_school_by_name(conn, school_name, school_cache)
        if school_id is None:
            school_not_found_warnings.append(
                f"{path.name}: no school found matching {school_name!r}. File skipped."
            )
            continue

        for row_idx, last_name, first_name, status, written_points, written_average in data_rows:
            if status is None and written_points is None and written_average is None:
                rows_skipped_blank += 1
                continue

            student_id, created = common.get_or_create_student(conn, first_name, last_name)
            if created:
                students_created += 1
            else:
                students_skipped += 1

            classroom_student_id, created = common.get_or_create_classroom_student(
                conn, classroom_id, student_id
            )
            if created:
                classroom_students_created += 1
            else:
                classroom_students_skipped += 1

            status_value = str(status).strip() if status is not None else None
            _, created = common.upsert_admission_written_result(
                conn, classroom_student_id, school_id, status_value, written_points, written_average
            )
            if created:
                admissions_created += 1
            else:
                admissions_updated += 1

    conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed.\n"
        f"Students created: {students_created}, skipped: {students_skipped}.\n"
        f"Classroom students created: {classroom_students_created}, skipped: {classroom_students_skipped}.\n"
        f"Admissions created: {admissions_created}, updated: {admissions_updated}.\n"
        f"Rows skipped (no data): {rows_skipped_blank}."
    )

    if school_not_found_warnings:
        print("\nWarnings:")
        for w in school_not_found_warnings:
            print(f"- {w}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("year", type=int)

    args = parser.parse_args()
    import_written_admissions(args.year)
