# AGENTS.md

## Overview

Builds per-student summaries of concours aux grandes écoles results for a classe prépa: raw `.xlsx` exports from `input/` are imported into SQLite (`db/concours_results.db`), then written to `output/Summary_{branch}_{year}.xlsx`.

Domain vocabulary used throughout the code and input files:
- **banque de concours / bank**: a group of schools sharing exams
- **écrits / written** and **admissibilité**: first stage. A student ends up `admissible` or `non admissible`.
- **oraux / oral** and **admission**: second stage. A student ends up `classé-e` (with a rank) or `non classé-e`.
- **filière / branch**: the classroom's track, one of `PC`, `MP`, `PSI` (`models/classroom.py`)
- **redoublant**: repeating student (5/2)

## Commands

See `README.md` for setup and running the whole pipeline.

Run a single step, as a module from the repo root so `config` and `models` resolve:
```bash
python -m db.init_db
python -m services.import_classroom_students 2026 input/students_2026.xlsx --branch PC
python -m services.import_written_admission_statuses 2026 --branch PC
python -m services.summarize_results 2026 --branch PC
```

Lint (CI runs both on pushes to `main` and on PRs). Ruff is a dev dependency, installed with `python -m pip install -e ".[dev]"`. `pyproject.toml` enables PEP 8 (`E`, `W`), naming (`N`) and a few docstring rules on top of ruff's defaults, with an 88-character line limit:
```bash
ruff check
ruff format --check
```

There is no test suite.

## Architecture

`main.py` runs an ordered pipeline for each branch. Each step is a `services/` module with a function of the same name and its own CLI. Module docstrings specify the exact expected file names, headers and behavior of each step, so read them before changing an importer.

The order matters because later steps depend on rows created by earlier ones:

| # | Step | Input folder | Creates or updates |
|---|------|--------------|--------------------|
| 1 | `import_classroom_students` | `input/students_{year}.xlsx` | classroom, students, classroom_students |
| 2 | `import_written_admission_statuses` | `input/ecrits/statuts_par_concours/` | **banks and schools**, admissions |
| 3 | `import_written_exam_results` | `input/ecrits/notes_par_concours/` | written exams, exam_results |
| 4 | `import_written_admission_scores` | `input/ecrits/notes_par_ecole/` | overwrites written fields on admissions |
| 5 | `import_oral_admission_statuses` | `input/oraux/statuts_par_concours/` | overwrites `oral_status` on admissions |
| 6 | `import_oral_exam_results_and_admissions` | `input/oraux/admissions_par_ecole/` | oral exams, exam_results, admission totals |
| 7 | `summarize_results` | database | output workbook, with `Admissions` and `Notes` tabs |

`summarize_results` runs after the written steps and again after the oral steps. `main.py` skips the oral steps, with a printed message, unless both oral input folders contain files matching the branch (`excel_utils.has_input_files`). A missing written input raises `FileNotFoundError` and aborts the run, so every branch being processed needs its written inputs present.

### Importer conventions

- **Validate, then write.** Every file is parsed first by `validate_and_read_file`, which returns an `excel_utils.ParsedFile` (`headers`, `rows`, `errors`, `is_empty`) whose rows are per-importer dataclasses. All format errors across all files are collected and raised as one `ValueError` by `excel_utils.raise_if_errors`, before the `with db_utils.transaction() as conn:` block opens.
- **Missing references warn rather than fail.** Only step 2 creates banks and schools. Later steps skip a file with a printed warning when its bank or school isn't found.
- **Name matching differs by entity.** Banks, schools and exams match loosely through `normalize_name` (in `models/text.py`, re-exported by `db_utils`) (ignores case, accents and punctuation). Students match exactly on `(first_name, last_name)`. The written steps create unknown students; the oral steps record an error and skip the row.
- **Statuses and exam formats are enums.** `models/admission.py` defines `ExamFormat` (`written`, `oral`) and `Status`, whose `Status.parse` maps the many spellings of admissible, non admissible, classé-e and non classé-e to one value. Both are `StrEnum`s, so the database still stores plain strings.
- **Re-running is expected.** The database persists between runs and every run re-imports everything. `get_or_create_*` helpers leave existing rows unchanged, so an `exam_results` row keeps its original points. `upsert_admission_written_result` and `update_admission` overwrite. Note that `update_admission` ignores `None` values, so it can't clear a field.
- **Schema changes need the database rebuilt.** `db/init_db.py` only uses `CREATE ... IF NOT EXISTS` and there are no migrations, so changes don't reach an existing database. School names are globally unique. Each table has an `updated_at` trigger generated from the `TABLES` list, so new tables must be added there too.

### Input files

- Files are selected by the per-branch `FilenamePattern`s in `excel_utils.filename_patterns`, and names that don't match are silently ignored. The pattern's regex captures the bank or school name from the file name. Names are NFC-normalized first because macOS can return decomposed accents.
- The students workbook has one sheet per branch. `validate_students_input` requires its sheets to exactly match the branches being run: all of `PC`, `MP` and `PSI` with no `--branch`, and only that branch's sheet with `--branch`.
- If several `students_{year}.xlsx` files exist, `main.py` prompts for the year on stdin.

### Personal data

`input/**/*.xlsx`, `db/concours_results.db` and `output/Summary*.xlsx` contain real student data and are gitignored, as is the `tmp/` scratch folder. Only `input/students_{year} (template).xlsx` is tracked.
