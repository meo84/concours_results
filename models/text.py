"""Text helpers shared by the models and the services."""

import unicodedata


def normalize_name(name: str) -> str:
    """Lowercase, strip accents and punctuation and collapse whitespace.

    Used for loose matching.
    """
    decomposed = unicodedata.normalize("NFKD", name.strip().lower())
    without_accents = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    cleaned = "".join(
        ch if ch.isalnum() or ch.isspace() else " " for ch in without_accents
    )
    return " ".join(cleaned.split())
