import openpyxl
from pathlib import Path
import unicodedata


def discover_files(folder: Path) -> list:
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(
            f"Missing input folder: {folder}. Please create it and add the results files."
        )
    files = sorted(f for f in folder.glob("*.xlsx") if not f.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(
            f"Input folder {folder} is empty. Please add the results files (.xlsx)."
        )
    return files


def read_xlsx_rows(path: Path) -> list[tuple]:
    """Read the first worksheet and return all rows."""
    wb = openpyxl.load_workbook(path, data_only=True)
    return list(wb.worksheets[0].iter_rows(values_only=True))


def validate_filename(
    path: Path,
    pattern: re.Pattern,
    name_group: str,
    expected_pattern: str,
):
    """Return (captured_name, error)."""
    normalized_filename = unicodedata.normalize("NFC", path.name)
    match = pattern.match(normalized_filename)

    if not match:
        return None, (
            f"{path.name}: filename does not match the expected pattern "
            f"'{expected_pattern}'. (raw: {normalized_filename!r})"
        )

    return match.group(name_group).strip(), None


def normalize_header_row(row) -> list[str | None]:
    """Strip trailing empty cells and normalize header values."""
    headers = list(row)

    while headers and headers[-1] is None:
        headers.pop()

    return [
        str(header).strip() if header is not None else None
        for header in headers
    ]

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
      return (f"{path_name}: file is empty")

    return None

