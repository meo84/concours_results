from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
DB_PATH = PROJECT_ROOT / "db" / "concours_results.db"
INPUT_PATH = PROJECT_ROOT / "input"
STUDENTS_PATH = INPUT_PATH / "students.xlsx"
WRITTEN_INPUT_PATH = INPUT_PATH / "ecrits"
WRITTEN_STATUSES_PER_BANK_PATH = WRITTEN_INPUT_PATH / "statuts_par_concours"
WRITTEN_EXAM_RESULTS_PATH  = WRITTEN_INPUT_PATH / "notes_par_concours"
WRITTEN_ADMISSIONS_PER_SCHOOL_PATH = WRITTEN_INPUT_PATH / "notes_par_ecole"
ORAL_ADMISSIONS_PER_SCHOOL_PATH = INPUT_PATH / "oraux" / "admissions_par_ecole"
OUTPUT_PATH = PROJECT_ROOT / "output" / "Summary.xlsx"
