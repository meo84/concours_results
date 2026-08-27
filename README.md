# Concours Results Summary

Builds a per-student summary of admission results for the concours
aux grandes écoles from a classe prépa.

The program:
- reads the raw results from `./input`
- imports them into an SQLite database
- generates a summary in `./output/Summary.xlsx`

## Setup

Requires Python 3.10 or later (see https://www.python.org/downloads/ if not installed)

Create and activate a virtual environment, then install the dependencies.

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e .
```

### Windows

```bash
py -m venv .venv
.venv\Scripts\activate
python -m pip install -e .
```

To inspect the SQLite database, you can install [DB Browser for SQLite](https://sqlitebrowser.org/dl/).

## Running the program

### Written results

When written results are published,
- fill out `input/students.xlsx` using `students_template.xlsx`
- download expected input files to `.input/ecrits/` and run:
```bash
python main.py
```

A summary is written to `.output/Summary.xlsx`.

### Oral results

When oral results are published,
add expected input files to `./input/oraux` and re-run:
```bash
python main.py
```

The summary is updated in `.output/Summary.xlsx`

## Deactivating the virtual environment

When you're done:

```bash
deactivate
```
