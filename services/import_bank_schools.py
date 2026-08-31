#!/usr/bin/env python3
"""
Import banks and schools from ./input/ecrits/statuts_par_concours/
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

Behavior:
    - Missing or empty folder -> raise FileNotFoundError.
    - Any filename or header format issue -> raise ValueError listing every
      issue found, across all files, before creating anything.
    - Valid folder ->
        - create a bank per distinct bank_name (loose match: same words,
          case-insensitive, accent- and punctuation-insensitive count as
          the same bank)
        - create a school per school-name column, scoped to its bank
          (same loose matching), only if it doesn't already exist

Usage:
    python -m services.import_banks_schools
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
    """Return (bank_name, school_names, errors).

    On any format error, returns (bank_name_or_None, None, errors).
    """
    bank_name, error = excel_utils.validate_filename(
        path, FILENAME_PATTERN, "bank_name",
        "Status pour l_admissibilité de la classe PC-PC pour la banque Banque {bank_name} PC...xlsx",
    )

    if error:
        return None, None, [error]

    all_rows = excel_utils.read_xlsx_rows(path)

    empty_file_error = excel_utils.validate_non_empty_file(all_rows, path.name)
    if empty_file_error:
        return bank_name, None, [empty_file_error]

    normalized_headers = excel_utils.normalize_header_row(all_rows[0])

    if len(normalized_headers) < 3 or normalized_headers[:3] != EXPECTED_FIRST_HEADERS:
        error = f"{path.name}: expected first 3 headers {EXPECTED_FIRST_HEADERS}, "
        f"found {normalized_headers[:3]}."
        return bank_name, None, [error]

    school_names = [h for h in normalized_headers[3:] if h]
    if not school_names:
        error = f"{path.name}: no school columns found after 'Prénom'."
        return bank_name, None, [error]

    return bank_name, school_names, []


def import_bank_schools() -> None:
    files = excel_utils.discover_files(WRITTEN_STATUSES_PER_BANK_PATH)

    all_errors = []
    parsed = []  # list of (path, bank_name, school_names)
    for path in files:
        bank_name, school_names, errors = validate_and_read_file(path)
        if errors:
            all_errors.extend(errors)
        else:
            parsed.append((path, bank_name, school_names))

    if all_errors:
        raise ValueError("Invalid input format:\n- " + "\n- ".join(all_errors))

    conn = db_utils.get_connection()

    try:
        with conn:
            bank_cache = {}
            school_cache = {}
            banks_created = 0
            schools_created = 0
            schools_skipped = 0

            for path, bank_name, school_names in parsed:
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

    finally:
        conn.close()

    print(
        f"\nDone. {len(parsed)} file(s) processed.\n"
        f"Banks created: {banks_created}.\n"
        f"Schools created: {schools_created}, skipped: {schools_skipped}.\n"
    )


if __name__ == "__main__":
    import_bank_schools()
