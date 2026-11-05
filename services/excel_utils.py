"""Helpers for locating, validating and reading input Excel files."""

import re
import unicodedata
from dataclasses import dataclass, field
from functools import cache
from pathlib import Path
from typing import Any

import openpyxl

from config import (
    INPUT_PATH,
    ORAL_ADMISSIONS_PER_SCHOOL_PATH,
    ORAL_STATUSES_PER_BANK_PATH,
    WRITTEN_ADMISSIONS_PER_SCHOOL_PATH,
    WRITTEN_EXAM_RESULTS_PATH,
    WRITTEN_STATUSES_PER_BANK_PATH,
)
from models.classroom import Branch

STUDENT_FILE_RE = re.compile(r"^students_(\d{4})\.xlsx$")


@dataclass(frozen=True)
class FilenamePattern:
    """How to recognize the input files of a folder and extract their name part."""

    pattern: re.Pattern
    group_name: str
    example: str


@dataclass(frozen=True)
class ParsedFile:
    """Result of validating and reading one input file.

    On format errors, `errors` is non-empty and `rows` is None. An empty file
    has `is_empty` set and neither headers, rows nor errors.
    """

    headers: list[str] | None = None
    rows: list[Any] | None = None
    errors: list[str] = field(default_factory=list)
    is_empty: bool = False


def validate_students_input(branches: list[Branch]) -> tuple[int, Path]:
    """Validate the students workbook and return (year, path).

    Looks for files named 'students_{year}.xlsx' in the inputs directory.
    Prompts the user if more than one is found.
    """
    matches = {}
    for f in sorted(INPUT_PATH.glob("students_*.xlsx")):
        m = STUDENT_FILE_RE.match(f.name)
        if m:
            year = int(m.group(1))
            matches[year] = f

    if not matches:
        raise FileNotFoundError(
            f"No 'students_{{year}}.xlsx' file found in '{INPUT_PATH}'."
        )

    if len(matches) == 1:
        [(year, students_path)] = matches.items()
    else:
        print("Multiple student files found:")
        for year in matches:
            print(f"  - {year}")
        choice = input("Type the year to use: ").strip()

        while not choice.isdigit() or int(choice) not in matches:
            choice = input("Not a valid year, try again: ").strip()

        year = int(choice)
        students_path = matches[year]

    wb = openpyxl.load_workbook(students_path, data_only=True)

    expected_sheets = set(branches)
    actual_sheets = set(wb.sheetnames)

    missing = expected_sheets - actual_sheets
    unexpected = actual_sheets - expected_sheets

    if missing or unexpected:
        raise ValueError(
            f"Invalid workbook sheets. "
            f"Missing: {sorted(missing)}, "
            f"Unexpected: {sorted(unexpected)}"
        )

    return year, students_path


def raise_if_errors(errors: list[str], title: str = "Invalid input format") -> None:
    """Raise a ValueError listing every error, if there are any."""
    if errors:
        raise ValueError(f"{title}:\n- " + "\n- ".join(errors))


def filter_files(folder: Path, branch: Branch) -> list[tuple[Path, str]]:
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(
            f"Missing input folder: {folder}. Please create it and add the results "
            "files."
        )
    expected_filename = filename_patterns(branch)[folder]
    filename_pattern = expected_filename.pattern
    group_name = expected_filename.group_name

    files = []
    for f in folder.iterdir():
        if not f.is_file() or f.name.startswith("~$"):
            continue

        normalized_name = unicodedata.normalize("NFC", f.name)
        match = filename_pattern.fullmatch(normalized_name)

        if match:
            files.append((f, match.group(group_name)))

    if not files:
        raise FileNotFoundError(
            f"Please add the results files to the input folder {folder}. "
            f"Expected file names: {expected_filename.example}"
        )

    return sorted(files)


def read_xlsx_rows(path: Path) -> list[tuple]:
    """Read the first worksheet and return all rows."""
    wb = openpyxl.load_workbook(path, data_only=True)
    return list(wb.worksheets[0].iter_rows(values_only=True))


def normalize_header_row(row) -> list[str | None]:
    """Strip trailing empty cells and normalize header values."""
    headers = list(row)

    while headers and headers[-1] is None:
        headers.pop()

    return [str(header).strip() if header is not None else None for header in headers]


def validate_name_columns(
    row_idx: int,
    last_name,
    first_name,
    filename: str,
) -> str | None:
    """Return an error for an incomplete name pair, otherwise None."""
    if last_name is None or first_name is None:
        return (
            f"{filename} row {row_idx}: missing Nom or Prénom "
            f"(Nom={last_name!r}, Prénom={first_name!r})."
        )

    return None


def validate_non_empty_file(rows: list[tuple], path_name: str) -> str | None:
    """Return an error when file is empty, otherwise None."""
    if not rows:
        return f"{path_name}: file is empty"

    return None


@cache
def filename_patterns(branch: Branch) -> dict[Path, FilenamePattern]:
    """Return the expected file name pattern of each input folder for a branch."""
    b = re.escape(branch)
    return {
        WRITTEN_STATUSES_PER_BANK_PATH: FilenamePattern(
            pattern=re.compile(
                rf"^Statuts\s+pour\s+l_admissibilité\s+de\s+la\s+classe\s+{b}-{b}"
                rf"\s+pour\s+la\s+banque\s+"
                rf"Banque\s+(?P<bank_name>.+)\s+{b}(?P<suffix>.*)\.xlsx$"
            ),
            group_name="bank_name",
            example=(
                f"Statuts pour l_admissibilité de la classe {b}-{b} "
                f"pour la banque Banque {{bank_name}} {b}...xlsx"
            ),
        ),
        WRITTEN_EXAM_RESULTS_PATH: FilenamePattern(
            pattern=re.compile(
                rf"^Résultats\s+de\s+la\s+classe\s+{b}-{b}\s+pour\s+la\s+Banque\s+"
                rf"(?P<bank_name>.+)\s+{b}(?P<suffix>.*)\.xlsx$"
            ),
            group_name="bank_name",
            example=(
                f"Résultats de la classe {b}-{b} "
                f"pour la Banque {{bank_name}} {b}...xlsx"
            ),
        ),
        WRITTEN_ADMISSIONS_PER_SCHOOL_PATH: FilenamePattern(
            pattern=re.compile(
                rf"^R[eé]sultats\s+de\s+l_admissibilit[eé]\s+pour\s+le\s+"
                rf"(?P<school_name>.+)\s+de\s+{b}(?P<suffix>.*)\.xlsx$"
            ),
            group_name="school_name",
            example=(
                f"Résultats de l_admissibilité pour le {{school_name}} de {b}...xlsx"
            ),
        ),
        ORAL_STATUSES_PER_BANK_PATH: FilenamePattern(
            pattern=re.compile(
                rf"^Statuts\s+pour\s+l_admission\s+de\s+la\s+classe\s+{b}-{b}"
                rf"\s+pour\s+la\s+banque\s+"
                rf"Banque\s+(?P<bank_name>.+)\s+{b}(?P<suffix>.*)\.xlsx$"
            ),
            group_name="bank_name",
            example=(
                f"Statuts pour l_admission de la classe {b}-{b} "
                f"pour la banque Banque {{bank_name}} {b}...xlsx"
            ),
        ),
        ORAL_ADMISSIONS_PER_SCHOOL_PATH: FilenamePattern(
            pattern=re.compile(
                rf"^R[eé]sultats\s+de\s+l_admission\s+pour\s+le\s+"
                rf"(?P<school_name>.+)\s+de\s+{b}(?P<suffix>.*)\.xlsx$"
            ),
            group_name="school_name",
            example=(f"Résultats de l_admission pour le {{school_name}} de {b}...xlsx"),
        ),
    }
