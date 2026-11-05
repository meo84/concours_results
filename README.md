# Concours Results Summary

Builds a per-student summary of admission results for the concours
aux grandes écoles from a classe prépa.

The program:
- reads the raw results from `./input`
- imports them into an SQLite database
- generates a summary per branch in `./output/Summary_{branch}_{year}.xlsx`

## Setup

Requires Python 3.11 or later (see https://www.python.org/downloads/ if not installed)

Create and activate a virtual environment, then install the dependencies.

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"
```

### Windows

```bash
py -m venv .venv
.venv\Scripts\activate
python -m pip install -e ".[dev]"
```

To inspect the SQLite database, you can install [DB Browser for SQLite](https://sqlitebrowser.org/dl/).

## Running the program

### Written results

When written results are published,
- fill out `input/students_{year}.xlsx` using `input/students_{year} (template).xlsx` - e.g. `students_2026.xlsx`
- download expected input files to `./input/ecrits/` and run:
```bash
python main.py
```

A summary per branch is written to `./output/Summary_{branch}_{year}.xlsx`.

### Running a single branch

By default, the program runs for all branches (`PC`, `MP` and `PSI`).
The students file must then have one sheet per branch, and the written results of every branch must be in `./input/ecrits/`.

To run for one branch only:
```bash
python main.py --branch PC
```

The students file must then have only that branch's sheet.

### Oral results

When oral results are published,
add expected input files to `./input/oraux` and re-run:
```bash
python main.py
```

The summaries are updated in `./output/Summary_{branch}_{year}.xlsx`.

## Deactivating the virtual environment

When you're done:

```bash
deactivate
```
