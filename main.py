import sys
from pathlib import Path

import argparse
import openpyxl

from config import (
    ORAL_ADMISSIONS_PER_SCHOOL_PATH,
    ORAL_STATUSES_PER_BANK_PATH,
    WRITTEN_ADMISSIONS_PER_SCHOOL_PATH,
    WRITTEN_EXAM_RESULTS_PATH,
    WRITTEN_STATUSES_PER_BANK_PATH,
)
from db.init_db import init_db
from models.classroom import Branch
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


def summarize_written_results(year: int, students_path: Path, branch: Branch) -> None:
    import_written_admission_statuses(year, branch)
    import_written_exam_results(year, branch)
    import_written_admission_scores(year, branch)
    summarize_results(year, branch)


def summarize_oral_results(year: int, branch: Branch) -> None:
    import_oral_admission_statuses(year, branch)
    import_oral_exam_results_and_admissions(year, branch)
    summarize_results(year, branch)


def process_results(year: int, students_path: Path, branch: Branch) -> None:
    import_classroom_students(year, students_path, branch)
    summarize_written_results(year, students_path, branch)
    try:
        summarize_oral_results(year, branch)
    except FileNotFoundError:
        print(f"No oral inputs found for {branch} {year}. Skipping oral results.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--branch",
        type=Branch,
        choices=list(Branch),
        metavar="{pc}",
    )
    args = parser.parse_args()
    branch = args.branch

    try:
        branches = [branch] if branch else list(Branch)
        year, students_path = excel_utils.validate_students_input(branches)
    except (FileNotFoundError, ValueError) as e:
        print(f"Students inputs are invalid: {e}")
        sys.exit(1)

    init_db()

    if branch:
        process_results(year, students_path, branch)
    else:
        for branch in Branch:
            process_results(year, students_path, branch)
