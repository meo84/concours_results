"""Import oral exam results and update admissions in concours_results.db.

Reads ./input/oraux/admissions_par_ecole/.

Expected folder content:
    - At least one .xlsx file.
    - Empty files are skipped silently (no error) — a file with no rows,
      or only a header row and no data rows, counts as empty.
    - For non-empty files, the filename follows the pattern:
        Resultats de l_admission pour le {school_name} de
        {branch}{optional more characters}.xlsx
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

Requires ./input/students_{year}.xlsx to have already been imported (via
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
                oral_status   <- Statut
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
    python -m services.import_oral_exam_results_and_admissions 2026 --branch PC
"""

import argparse
from dataclasses import dataclass
from pathlib import Path

from config import ORAL_ADMISSIONS_PER_SCHOOL_PATH
from models.classroom import Branch
from services import db_utils, excel_utils

EXPECTED_HEADERS = [
    "Numéro",
    "Nom",
    "Prénom",
    "Statut",
    "Rang",
    "Total",
    "Moyenne",
    "Total oral",
    "total écrit",
]
N_VALIDATED_COLUMNS = len(EXPECTED_HEADERS)  # 9

FORMAT = "oral"


@dataclass(frozen=True)
class OralRow:
    """A student row of a per-school oral results file.

    extra_values is aligned with the file's exam columns (None for a blank cell).
    """

    row_idx: int
    last_name: str
    first_name: str
    oral_status: str | None
    rank: int | None
    total: float | None
    average: float | None
    total_oral: float | None
    extra_values: list


@dataclass(frozen=True)
class ResolvedOralRow:
    """An OralRow whose student is part of the classroom."""

    row: OralRow
    classroom_student_id: int


def _is_effectively_empty(all_rows) -> bool:
    """Return True if the file has no usable content.

    That is: no rows, only a header row (no data rows), or rows/columns
    present but every cell (including headers) blank.
    """
    if len(all_rows) <= 1:
        return True  # no rows, or header row only
    all_cells = [v for row in all_rows for v in row]
    return all(v is None or (isinstance(v, str) and v.strip() == "") for v in all_cells)


def validate_and_read_file(path: Path) -> excel_utils.ParsedFile:
    """Return a ParsedFile with the exam column names as headers.

    rows is a list of OralRow whose extra_values are aligned with the headers.
    On any format error, rows is None. A genuinely empty file has is_empty set.
    """
    all_rows = excel_utils.read_xlsx_rows(path)

    if _is_effectively_empty(all_rows):
        return excel_utils.ParsedFile(is_empty=True)

    errors = []

    normalized_headers = excel_utils.normalize_header_row(all_rows[0])

    if len(normalized_headers) < N_VALIDATED_COLUMNS:
        errors.append(
            f"{path.name}: expected at least {N_VALIDATED_COLUMNS} columns starting "
            "with "
            f"{EXPECTED_HEADERS}, found {normalized_headers}."
        )
        return excel_utils.ParsedFile(errors=errors)

    actual_headers = normalized_headers[:N_VALIDATED_COLUMNS]
    for expected, actual in zip(EXPECTED_HEADERS, actual_headers):
        if db_utils.normalize_name(actual or "") != db_utils.normalize_name(expected):
            errors.append(
                f"{path.name}: expected header {expected!r} (loose match), found "
                f"{actual!r}."
            )
    if errors:
        return excel_utils.ParsedFile(errors=errors)

    extra_headers = [h for h in normalized_headers[N_VALIDATED_COLUMNS:] if h]
    n_extra = len(extra_headers)

    data_rows = []
    for row_idx, row in enumerate(all_rows[1:], start=2):
        last_name = row[1] if len(row) > 1 else None
        first_name = row[2] if len(row) > 2 else None
        oral_status = row[3] if len(row) > 3 else None
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
        name_columns_error = excel_utils.validate_name_columns(
            row_idx, last_name, first_name, path.name
        )
        if name_columns_error:
            errors.append(name_columns_error)
            continue

        extra_values = list(row[N_VALIDATED_COLUMNS : N_VALIDATED_COLUMNS + n_extra])
        extra_values += [None] * max(0, n_extra - len(extra_values))

        data_rows.append(
            OralRow(
                row_idx,
                str(last_name).strip(),
                str(first_name).strip(),
                oral_status,
                rank,
                total,
                average,
                total_oral,
                extra_values,
            )
        )

    if errors:
        return excel_utils.ParsedFile(headers=extra_headers, errors=errors)

    return excel_utils.ParsedFile(headers=extra_headers, rows=data_rows)


def import_oral_exam_results_and_admissions(year: int, branch: Branch) -> None:
    files = excel_utils.filter_files(ORAL_ADMISSIONS_PER_SCHOOL_PATH, branch)

    all_errors = []
    parsed = []  # list of (path, school_name, extra_headers, data_rows)
    skipped_empty = 0
    for path, school_name in files:
        result = validate_and_read_file(path)
        if result.is_empty:
            skipped_empty += 1
            continue
        if result.errors:
            all_errors.extend(result.errors)
        else:
            parsed.append((path, school_name, result.headers, result.rows))

    if all_errors:
        raise ValueError("Invalid input format:\n- " + "\n- ".join(all_errors))

    conn = db_utils.get_connection()

    try:
        with conn:
            classroom_id = db_utils.get_classroom_id_by_year(conn, year, branch)
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
                school_id = db_utils.find_school_by_name(
                    conn, school_name, school_cache
                )
                if school_id is None:
                    school_not_found_warnings.append(
                        f"{path.name}: no school found matching {school_name!r}. File "
                        "skipped."
                    )
                    continue

                bank_id = db_utils.get_school_bank_id(conn, school_id)

                # Resolve classroom_student_id per row; rows that fail are excluded
                # from both the admissions update and the exam results creation.
                resolved_rows: list[ResolvedOralRow] = []
                for row in data_rows:
                    student_id = db_utils.find_student_by_name(
                        conn, row.first_name, row.last_name
                    )
                    if student_id is None:
                        row_errors.append(
                            f"{path.name} row {row.row_idx}: no student found for "
                            f"{row.first_name!r} {row.last_name!r}. Row skipped."
                        )
                        continue

                    classroom_student_id = db_utils.find_classroom_student(
                        conn, classroom_id, student_id
                    )
                    if classroom_student_id is None:
                        row_errors.append(
                            f"{path.name} row {row.row_idx}: no classroom_student "
                            "found for "
                            f"{row.first_name!r} {row.last_name!r} in classroom "
                            f"{classroom_id}. Row skipped."
                        )
                        continue

                    resolved_rows.append(ResolvedOralRow(row, classroom_student_id))

                    admission = db_utils.find_admission(
                        conn, classroom_student_id, school_id
                    )
                    if admission is None:
                        row_errors.append(
                            f"{path.name} row {row.row_idx}: no admissions record "
                            "found for "
                            f"{row.first_name!r} {row.last_name!r} at "
                            f"school_id={school_id}. "
                            f"Admissions update skipped for this row."
                        )
                        continue

                    admission_id, written_points = admission
                    if row.total_oral is not None:
                        oral_points = row.total_oral
                    elif row.total is not None and written_points is not None:
                        oral_points = row.total - written_points
                    else:
                        oral_points = None

                    db_utils.update_admission(
                        conn,
                        admission_id,
                        oral_status=row.oral_status,
                        rank=row.rank,
                        total_points=row.total,
                        average=row.average,
                        oral_points=oral_points,
                    )
                    admissions_updated += 1

                # Per-exam-column processing.
                for col_idx, exam_name in enumerate(extra_headers):
                    column_values = [r.row.extra_values[col_idx] for r in resolved_rows]
                    if all(v is None for v in column_values):
                        continue  # column entirely blank across resolved rows -> skip

                    exam_id, created = db_utils.get_or_create_exam(
                        conn, exam_name, bank_id, FORMAT, exam_cache
                    )
                    if created:
                        exams_created += 1
                        print(
                            f"Created exam: {exam_name!r} (id={exam_id}, "
                            f"bank_id={bank_id}, format={FORMAT})"
                        )
                    else:
                        exams_skipped += 1

                    for resolved in resolved_rows:
                        points = resolved.row.extra_values[col_idx]
                        if points is None:
                            continue
                        _, created = db_utils.get_or_create_exam_result(
                            conn, resolved.classroom_student_id, exam_id, points
                        )
                        if created:
                            results_created += 1
                        else:
                            results_skipped += 1

    finally:
        conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed, {skipped_empty} empty file(s) "
        "skipped.\n"
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
    parser = argparse.ArgumentParser()
    parser.add_argument("year", type=int)
    parser.add_argument(
        "--branch",
        type=Branch,
        choices=list(Branch),
    )

    args = parser.parse_args()
    branch = args.branch
    if branch:
        import_oral_exam_results_and_admissions(args.year, branch)
    else:
        for branch in Branch:
            import_oral_exam_results_and_admissions(args.year, branch)
