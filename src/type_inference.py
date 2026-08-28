"""
type_inference.py
------------------
Domain-agnostic column type inference.

Given a pandas Series of raw (string) values, guesses which semantic type
the column represents by scoring it against a battery of pattern
detectors, then returns the type whose detector matched the highest share
of non-null values. This is the piece that lets the profiler work on ANY
spreadsheet — an HR export, a project tracker, a sales ledger — without
being told anything about the schema up front.

Design note: every detector returns a MATCH RATE (0-1), not a boolean, so
a column that's "90% clean emails, 10% garbage" is still correctly typed
as `email` — the 10% becomes the validity problem the profiler reports,
not a reason to give up and call the column free text.
"""

import re
from dataclasses import dataclass

import numpy as np
import pandas as pd

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$")
URL_RE = re.compile(r"^(https?://|www\.)[^\s]+\.[a-zA-Z]{2,}[^\s]*$", re.IGNORECASE)
PHONE_RE = re.compile(r"^\+?\d[\d\s\-()]{6,15}\d$")
SAUDI_ID_RE = re.compile(r"^[12]\d{9}$")
INTEGER_RE = re.compile(r"^-?\d+$")
FLOAT_RE = re.compile(r"^-?\d+\.\d+$")
CURRENCY_RE = re.compile(r"^-?\d[\d,]*(\.\d+)?\s?(SAR|USD|\$|ر\.س)?$", re.IGNORECASE)
BOOLEAN_VALUES = {"true", "false", "yes", "no", "y", "n", "1", "0", "نعم", "لا"}
PERCENT_RE = re.compile(r"^-?\d+(\.\d+)?\s?%$")

DATE_FORMATS = [
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%y",
    "%Y/%m/%d", "%d %b %Y", "%b %d, %Y", "%d.%m.%Y",
]

# Order matters as a tie-breaker when two detectors score identically:
# dates must be resolved before phone (an ISO date like "2024-03-15" is
# digits-and-separators and would otherwise satisfy a loose phone regex).
TYPE_ORDER = [
    "email", "url", "saudi_id", "date", "phone", "boolean", "percent",
    "integer", "float", "currency", "categorical", "free_text",
]


def _match_rate(series: pd.Series, predicate) -> float:
    non_empty = series[series.astype(str).str.strip() != ""]
    if len(non_empty) == 0:
        return 0.0
    hits = non_empty.astype(str).str.strip().apply(predicate).sum()
    return hits / len(non_empty)


def _is_date(v: str) -> bool:
    for fmt in DATE_FORMATS:
        try:
            pd.to_datetime(v, format=fmt)
            return True
        except (ValueError, TypeError):
            continue
    return False


def _is_categorical(series: pd.Series) -> float:
    non_empty = series[series.astype(str).str.strip() != ""]
    if len(non_empty) == 0:
        return 0.0
    n_unique = non_empty.nunique()
    # Heuristic: few distinct values relative to row count, repeated often,
    # and none of them individually look numeric/date/email-like.
    ratio_unique = n_unique / len(non_empty)
    avg_len = non_empty.astype(str).str.len().mean()
    if ratio_unique <= 0.5 and n_unique <= 50 and avg_len <= 40:
        return 1.0 - ratio_unique
    return 0.0


DETECTORS = {
    "email": lambda s: _match_rate(s, lambda v: bool(EMAIL_RE.match(v))),
    "url": lambda s: _match_rate(s, lambda v: bool(URL_RE.match(v))),
    "saudi_id": lambda s: _match_rate(s, lambda v: bool(SAUDI_ID_RE.match(v))),
    "phone": lambda s: _match_rate(s, lambda v: bool(PHONE_RE.match(v)) and not _is_date(v)),
    "boolean": lambda s: _match_rate(s, lambda v: v.lower() in BOOLEAN_VALUES),
    "percent": lambda s: _match_rate(s, lambda v: bool(PERCENT_RE.match(v))),
    "date": lambda s: _match_rate(s, _is_date),
    "integer": lambda s: _match_rate(s, lambda v: bool(INTEGER_RE.match(v))),
    "float": lambda s: _match_rate(s, lambda v: bool(FLOAT_RE.match(v))),
    "currency": lambda s: _match_rate(s, lambda v: bool(CURRENCY_RE.match(v)) and any(c.isdigit() for c in v)),
    "categorical": _is_categorical,
    "free_text": lambda s: 0.15,  # low-confidence fallback, always available
}

MIN_CONFIDENCE = 0.6  # below this, we fall back to free_text


@dataclass
class ColumnTypeGuess:
    column: str
    inferred_type: str
    confidence: float
    scores: dict


def infer_column_type(series: pd.Series, column_name: str = "") -> ColumnTypeGuess:
    scores = {}
    for t in TYPE_ORDER:
        try:
            scores[t] = round(float(DETECTORS[t](series)), 3)
        except Exception:
            scores[t] = 0.0

    # column-name hints nudge close calls (e.g. "id" columns that are
    # purely numeric shouldn't be typed as free-flowing integers of no
    # semantic meaning — still reported as `integer`/`id-like`, informational only)
    best_type = max(scores, key=scores.get)
    best_score = scores[best_type]

    if best_score < MIN_CONFIDENCE and best_type != "categorical":
        best_type = "free_text"
        best_score = scores["free_text"]

    return ColumnTypeGuess(column=column_name, inferred_type=best_type, confidence=best_score, scores=scores)
