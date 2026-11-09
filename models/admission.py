"""Admission domain types: exam formats and admission statuses."""

from enum import StrEnum

from models.text import normalize_name


class ExamFormat(StrEnum):
    """Stage of the concours an exam belongs to, as stored in `exams.format`."""

    WRITTEN = "written"
    ORAL = "oral"


class Status(StrEnum):
    """Admission outcome: admissibilité (written) or admission (oral)."""

    ADMISSIBLE = "admissible"
    NON_ADMISSIBLE = "non admissible"
    CLASSE = "classé-e"
    NON_CLASSE = "non classé-e"

    @classmethod
    def parse(cls, raw: str) -> "Status | None":
        """Return the status a raw cell value stands for, or None if unknown.

        Matching ignores case, accents and punctuation, so "Classé-e",
        "classee" and "classée" are all `Status.CLASSE`.
        """
        return _STATUS_BY_NORMALIZED_NAME.get(normalize_name(raw))


_STATUS_BY_NORMALIZED_NAME = {
    normalize_name(alias): status
    for status, aliases in {
        Status.ADMISSIBLE: ["admissible"],
        Status.NON_ADMISSIBLE: ["non admissible"],
        Status.CLASSE: ["classé-e", "classee", "classée", "classe-e"],
        Status.NON_CLASSE: [
            "non classé-e",
            "non classee",
            "non classée",
            "non classe-e",
        ],
    }.items()
    for alias in aliases
}
