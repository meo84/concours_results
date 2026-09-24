import sys
from pathlib import Path

import openpyxl

from config import (
    ORAL_ADMISSIONS_PER_SCHOOL_PATH,
    ORAL_STATUSES_PER_BANK_PATH,
    WRITTEN_ADMISSIONS_PER_SCHOOL_PATH,
    WRITTEN_EXAM_RESULTS_PATH,
    WRITTEN_STATUSES_PER_BANK_PATH,
)
from db.init_db import init_db
from services import excel_utils
from services.import_classroom_students import import_classroom_students
from services.import_oral_admission_statuses import import_oral_admission_statuses
from services.import_oral_exam_results_and_admissions import (
    import_oral_exam_results_and_admissions,
)
from services.import_written_admission_scores import import_written_admission_scores
from services.import_written_admission_statuses import import_written_admission_statuses
from services.import_written_exam_results import import_written_exam_results
from services.summarize_results import summarize_results


def validate_written_inputs() -> None:
    excel_utils.discover_files(WRITTEN_STATUSES_PER_BANK_PATH)
    excel_utils.discover_files(WRITTEN_EXAM_RESULTS_PATH)
    excel_utils.discover_files(WRITTEN_ADMISSIONS_PER_SCHOOL_PATH)


def summarize_written_results(year: int, students_path: Path) -> None:
    import_classroom_students(year, students_path)
    import_written_admission_statuses(year)
    import_written_exam_results(year)
    import_written_admission_scores(year)
    summarize_results(year)


def validate_oral_inputs() -> None:
    excel_utils.discover_files(ORAL_STATUSES_PER_BANK_PATH)
    excel_utils.discover_files(ORAL_ADMISSIONS_PER_SCHOOL_PATH)


def summarize_oral_results(year: int) -> None:
    import_oral_admission_statuses(year)
    import_oral_exam_results_and_admissions(year)
    summarize_results(year)


if __name__ == "__main__":
    try:
        year, students_path = excel_utils.validate_students_input()
        validate_written_inputs()
    except FileNotFoundError as e:
        print(f"Written inputs are invalid: {e}")
        sys.exit(1)

    init_db()
    summarize_written_results(year, students_path)

    try:
        validate_oral_inputs()
    except FileNotFoundError:
        print("No oral inputs found. Skipping oral results.")
    else:
        summarize_oral_results(year)
