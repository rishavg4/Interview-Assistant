"""
cache_service.py
Tracks interview session counts per role and performs fuzzy question matching
against the Excel history using rapidfuzz for fast similarity scoring.

Search strategy:
  1. Search the selected role's sheet first (primary).
  2. If no match, search every other role's sheet in order (cross-role fallback).
  3. Return the best match found, along with which role it came from.
  4. If nothing matches anywhere, return None → AI is called.
"""

import json
from pathlib import Path
from rapidfuzz import process, fuzz

from services import excel_service

DATA_DIR = Path(__file__).resolve().parents[2] / "data"
STATS_FILE = DATA_DIR / "stats.json"

# Similarity score threshold (0–100). Higher = stricter matching.
SIMILARITY_THRESHOLD = 65


# ── Stats helpers ──────────────────────────────────────────────────────────

def _read_stats() -> dict:
    if not STATS_FILE.exists():
        return {}
    try:
        return json.loads(STATS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _write_stats(stats: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    STATS_FILE.write_text(json.dumps(stats, indent=2), encoding="utf-8")


def get_stats() -> dict:
    return _read_stats()


def get_interview_count(role: str) -> int:
    return _read_stats().get(role, 0)


def increment_interview_count(role: str) -> None:
    stats = _read_stats()
    stats[role] = stats.get(role, 0) + 1
    _write_stats(stats)


def get_cache_size(role: str) -> int:
    """Return number of saved Q&A rows for a role."""
    return len(excel_service.get_history(role))


# ── Fuzzy match helpers ────────────────────────────────────────────────────

def _search_in_role(question: str, role: str) -> dict | None:
    """
    Search one role's sheet for a fuzzy match.
    Returns the matched row dict (includes 'role' key) or None.
    """
    history = excel_service.get_history(role)
    if not history:
        return None

    questions = [row["question"] for row in history]
    result = process.extractOne(
        question,
        questions,
        scorer=fuzz.token_set_ratio,
        score_cutoff=SIMILARITY_THRESHOLD,
    )

    if result is None:
        return None

    _matched_q, _score, index = result
    best = history[index]

    if not best.get("points"):
        return None

    return best


# ── Public API ─────────────────────────────────────────────────────────────

def find_similar_answer(role: str, question: str) -> dict | None:
    """
    Search for a question similar to the given one.

    Strategy:
      1. Search the selected role's sheet (primary).
      2. If not found, search all other role sheets in order (cross-role fallback).

    Returns:
      {
        'points':       [...],
        'illustrated':  '...',
        'matched_role': '<role name>'   # which sheet the answer came from
      }
      or None if nothing matched anywhere.
    """
    # ── Step 1: Primary role search ────────────────────────────────────────
    match = _search_in_role(question, role)
    if match:
        return {
            "points":       match["points"],
            "illustrated":  match["illustrated"],
            "matched_role": role,
        }

    # ── Step 2: Cross-role fallback ────────────────────────────────────────
    other_roles = [r for r in excel_service.ROLES if r != role]
    for other_role in other_roles:
        match = _search_in_role(question, other_role)
        if match:
            return {
                "points":       match["points"],
                "illustrated":  match["illustrated"],
                "matched_role": match.get("role") or other_role,
            }

    return None
