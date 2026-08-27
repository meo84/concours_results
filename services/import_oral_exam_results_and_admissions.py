#!/usr/bin/env python3
"""
Import oral exam results and update admissions from
./input/oraux/admissions_par_ecole/ into concours_results.db.

Expected folder content:
    - At least one .xlsx file.
    - Empty files are skipped silently (no error) — a file with no rows,
      or only a header row and no data rows, counts as empty.
    - For non-empty files, the filename follows the pattern:
        Resultats de l_admission pour le {school_name} de PC{optional more characters}.xlsx
      e.g. "Resultats de l_admission pour le Concours Mines-Télécom
      fonctionnaire PC de PC.xlsx"
      -> school_name = "Mines-Télécom fonctionnaire PC"
      (the accent on the first word is tolerated either way: "Résultats" or
      "Resultats", since both appear in practice)
    - In each non-empty file, the first 9 columns must be, in order:
        Numéro, Nom, Prénom, Statut, Rang, Total, Moyenne, Total oral,
        total écrit
      All 9 headers use loose matching (accent/punctuation/
      case-insensitive). Columns after these 9 are per-exam oral result
      columns (dynamic, one per oral exam). Extra columns beyond the 9 are
      allowed.

Requires ./input/students.xlsx to have already been imported (via
import_classroom_students.py)

Behavior:
    - Missing or empty folder -> raise FileNotFoundError.
    - Any filename or header format issue (in a non-empty file) -> raise
      ValueError listing every issue found, across all files, before
      creating anything.
    - No classroom found for the given year -> raise ValueError
      telling you to run import_classroom_students.py first.
    - Valid folder ->
        - for each non-empty file, find the school by name (loose match).
          If no matching school exists, the file is skipped and a warning
          is recorded (printed in the summary at the end) — processing
          continues with the remaining files.
        - for each data row:
            - find the student (exact match on first_name/last_name) and
              the classroom_student (exact match on classroom_id,
              student_id). Neither is created here — both are expected to
              already exist from earlier imports. If either is missing,
              the row is skipped and an error is recorded (printed at the
              end) — processing continues with the remaining rows.
            - find the admissions record (classroom_student_id, school_id).
              If missing, the admissions update for this row is skipped and
              an error is recorded — but the row is still eligible for the
              per-exam result columns below, since those don't require an
              admissions row. If found, it is updated (overwritten, not
              just filled in) with:
                status        <- Statut
                rank          <- Rang
                total_points  <- Total
                average       <- Moyenne
                oral_points   <- Total oral if provided, else
                                 total_points (this file's Total) minus the
                                 admission's existing written_points
        - for each column after the 9 validated columns: if every row in
          the file has a blank cell in that column, it's skipped entirely
          (no exam is created for it). Otherwise:
            - find or create an exam (format="oral", bank_id from the
              file's school, name = the column header) if one doesn't
              already exist for that (bank_id, name, format) combination
              (loose name matching)
            - for each row with a non-blank cell in that column (and a
              resolved classroom_student), find or create an exam_results
              record (classroom_student_id, exam_id, points)

Usage:
    python import_oral_exam_results_and_admissions.py
"""

import argparse
import re
import unicodedata
from pathlib import Path

import openpyxl

from services import common
from services import import_classroom_students
from config import ORAL_ADMISSIONS_PER_SCHOOL_PATH

EXPECTED_HEADERS = ["Numéro", "Nom", "Prénom", "Statut", "Rang", "Total", "Moyenne", "Total oral", "total écrit"]
N_VALIDATED_COLUMNS = len(EXPECTED_HEADERS)  # 9

FORMAT = "oral"

FILENAME_PATTERN = re.compile(
    r"^R[eé]sultats\s+de\s+l_admission\s+pour\s+le\s+"
    r"(?P<school_name>.+)\s+de\s+PC(?P<suffix>.*)\.xlsx$"
)


def _is_effectively_empty(all_rows) -> bool:
    """True if the file has no rows, only a header row (no data rows), or
    has rows/columns present but every cell (including headers) is blank."""
    if len(all_rows) <= 1:
        return True  # no rows, or header row only
    all_cells = [v for row in all_rows for v in row]
    return all(v is None or (isinstance(v, str) and v.strip() == "") for v in all_cells)


def validate_and_read_file(path: Path):
    """Return (school_name, extra_headers, data_rows, errors, is_empty).

    data_rows is a list of (row_idx, last_name, first_name, status, rank,
    total, average, total_oral, extra_values) where extra_values is a list
    aligned with extra_headers (None for a blank cell). On any format
    error, returns (school_name_or_None, None, None, errors, False).
    A genuinely empty file returns (None, None, None, [], True).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))

    if _is_effectively_empty(all_rows):
        return None, None, None, [], True

    errors = []

    normalized_filename = unicodedata.normalize("NFC", path.name)
    match = FILENAME_PATTERN.match(normalized_filename)
    if not match:
        errors.append(
            f"{path.name}: filename does not match the expected pattern "
            f"'Résultats de l_admission pour le {{school_name}} de PC...xlsx'. "
            f"(raw: {normalized_filename!r})"
        )
        return None, None, None, errors, False

    school_name = match.group("school_name").strip()

    header_row = list(all_rows[0])
    while header_row and header_row[-1] is None:
        header_row.pop()
    normalized_headers = [str(h).strip() if h is not None else None for h in header_row]

    if len(normalized_headers) < N_VALIDATED_COLUMNS:
        errors.append(
            f"{path.name}: expected at least {N_VALIDATED_COLUMNS} columns starting with "
            f"{EXPECTED_HEADERS}, found {normalized_headers}."
        )
        return school_name, None, None, errors, False

    actual_headers = normalized_headers[:N_VALIDATED_COLUMNS]
    for expected, actual in zip(EXPECTED_HEADERS, actual_headers):
        if common.normalize_name(actual or "") != common.normalize_name(expected):
            errors.append(
                f"{path.name}: expected header {expected!r} (loose match), found {actual!r}."
            )
    if errors:
        return school_name, None, None, errors, False

    extra_headers = [h for h in normalized_headers[N_VALIDATED_COLUMNS:] if h]
    n_extra = len(extra_headers)

    data_rows = []
    for row_idx, row in enumerate(all_rows[1:], start=2):
        last_name = row[1] if len(row) > 1 else None
        first_name = row[2] if len(row) > 2 else None
        status = row[3] if len(row) > 3 else None
        rank = row[4] if len(row) > 4 else None
        total = row[5] if len(row) > 5 else None
        average = row[6] if len(row) > 6 else None
        total_oral = row[7] if len(row) > 7 else None
        # row[8] is 'total écrit' — validated but its value isn't read here.
        # The oral_points fallback formula uses the admissions row's
        # existing written_points (set earlier by import_admissions.py),
        # which is the same value this column would contain — so reading
        # it again from this file would be redundant, not missing.

        if last_name is None and first_name is None:
            continue  # skip fully blank rows
        if last_name is None or first_name is None:
            errors.append(
                f"{path.name} row {row_idx}: missing Nom or Prénom "
                f"(Nom={last_name!r}, Prénom={first_name!r})."
            )
            continue

        extra_values = list(row[N_VALIDATED_COLUMNS:N_VALIDATED_COLUMNS + n_extra])
        extra_values += [None] * max(0, n_extra - len(extra_values))

        data_rows.append((
            row_idx,
            str(last_name).strip(),
            str(first_name).strip(),
            status,
            rank,
            total,
            average,
            total_oral,
            extra_values,
        ))

    if errors:
        return school_name, extra_headers, None, errors, False

    return school_name, extra_headers, data_rows, errors, False


def import_oral_exam_results_and_admissions(year: int) -> None:
    files = common.discover_files(ORAL_ADMISSIONS_PER_SCHOOL_PATH)

    all_errors = []
    parsed = []  # list of (path, school_name, extra_headers, data_rows)
    skipped_empty = 0
    for path in files:
        school_name, extra_headers, data_rows, errors, is_empty = validate_and_read_file(path)
        if is_empty:
            skipped_empty += 1
            continue
        if errors:
            all_errors.extend(errors)
        else:
            parsed.append((path, school_name, extra_headers, data_rows))

    if all_errors:
        raise ValueError("Invalid input format:\n- " + "\n- ".join(all_errors))

    conn = common.get_connection()

    classroom_id = common.get_classroom_id_by_year(conn, year)
    print(f"Using classroom for year {year} (id={classroom_id})")

    school_cache = {}
    exam_cache = {}
    school_not_found_warnings = []
    row_errors = []
    admissions_updated = 0
    exams_created = 0
    exams_skipped = 0
    results_created = 0
    results_skipped = 0

    for path, school_name, extra_headers, data_rows in parsed:
        school_id = common.find_school_by_name(conn, school_name, school_cache)
        if school_id is None:
            school_not_found_warnings.append(
                f"{path.name}: no school found matching {school_name!r}. File skipped."
            )
            continue

        bank_id = common.get_school_bank_id(conn, school_id)

        # Resolve classroom_student_id per row; rows that fail are excluded
        # from both the admissions update and the exam results creation.
        resolved_rows = []  # list of (row_idx, classroom_student_id, status, rank, total, average, total_oral, extra_values)
        for row_idx, last_name, first_name, status, rank, total, average, total_oral, extra_values in data_rows:
            student_id = common.find_student_by_name(conn, first_name, last_name)
            if student_id is None:
                row_errors.append(
                    f"{path.name} row {row_idx}: no student found for "
                    f"{first_name!r} {last_name!r}. Row skipped."
                )
                continue

            classroom_student_id = common.find_classroom_student(conn, classroom_id, student_id)
            if classroom_student_id is None:
                row_errors.append(
                    f"{path.name} row {row_idx}: no classroom_student found for "
                    f"{first_name!r} {last_name!r} in classroom {classroom_id}. Row skipped."
                )
                continue

            resolved_rows.append(
                (row_idx, classroom_student_id, status, rank, total, average, total_oral, extra_values)
            )

            admission = common.find_admission(conn, classroom_student_id, school_id)
            if admission is None:
                row_errors.append(
                    f"{path.name} row {row_idx}: no admissions record found for "
                    f"{first_name!r} {last_name!r} at school_id={school_id}. "
                    f"Admissions update skipped for this row."
                )
                continue

            admission_id, written_points = admission
            if total_oral is not None:
                oral_points = total_oral
            elif total is not None and written_points is not None:
                oral_points = total - written_points
            else:
                oral_points = None

            common.update_admission_oral_result(
                conn, admission_id, status, rank, total, average, oral_points
            )
            admissions_updated += 1

        # Per-exam-column processing.
        for col_idx, exam_name in enumerate(extra_headers):
            column_values = [r[7][col_idx] for r in resolved_rows]
            if all(v is None for v in column_values):
                continue  # column entirely blank across resolved rows -> skip

            exam_id, created = common.get_or_create_exam(conn, exam_name, bank_id, FORMAT, exam_cache)
            if created:
                exams_created += 1
                print(f"Created exam: {exam_name!r} (id={exam_id}, bank_id={bank_id}, format={FORMAT})")
            else:
                exams_skipped += 1

            for row_idx, classroom_student_id, *_rest, extra_values in resolved_rows:
                points = extra_values[col_idx]
                if points is None:
                    continue
                _, created = common.get_or_create_exam_result(
                    conn, classroom_student_id, exam_id, points
                )
                if created:
                    results_created += 1
                else:
                    results_skipped += 1

    conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed, {skipped_empty} empty file(s) skipped.\n"
        f"Admissions updated: {admissions_updated}.\n"
        f"Exams created: {exams_created}, skipped: {exams_skipped}.\n"
        f"Exam results created: {results_created}, skipped: {results_skipped}."
    )

    if school_not_found_warnings:
        print("\nSchool warnings:")
        for w in school_not_found_warnings:
            print(f"- {w}")

    if row_errors:
        print("\nRow errors:")
        for e in row_errors:
            print(f"- {e}")


if __name__ == "__main__":
    main()
    parser = argparse.ArgumentParser()
    parser.add_argument("year", type=int)

    args = parser.parse_args()
    import_oral_exam_results_and_admissions(args.year)
