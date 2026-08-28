"""
profiler.py
-----------
Domain-agnostic data quality profiler.

Given ANY tabular dataset (DataFrame loaded from .xlsx/.csv), this module:
  1. Infers the semantic type of every column (type_inference.py)
  2. Scores completeness (non-missing %), validity (conforms to the
     inferred type's pattern), and uniqueness for every column
  3. Flags outliers for numeric columns (IQR method) and rare/one-off
     categories for categorical columns
  4. Flags exact duplicate rows at the dataset level
  5. Rolls everything up into a 0-100 quality score per column and overall

No column name or column count is hardcoded anywhere in this file — the
same code profiles an HR export, a project tracker, or a sales ledger.
"""

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from collections import Counter

from src.type_inference import _is_date, infer_column_type

NUMERIC_TYPES = {"integer", "float", "currency", "percent"}
# columns where near-uniqueness is actually part of "correctness" (an ID-like
# column full of repeats is a quality problem); categorical/date/numeric
# columns are expected to repeat values, so uniqueness isn't a fair dimension
# to judge them on.
UNIQUENESS_RELEVANT_TYPES = {"email", "saudi_id", "phone", "url"}
DIMENSION_LABELS = ("Completeness", "Validity", "Uniqueness")


@dataclass
class ColumnProfile:
    column: str
    inferred_type: str
    confidence: float
    completeness_pct: float
    validity_pct: float
    uniqueness_pct: float
    issue_count: int
    quality_score: float
    sample_invalid: list = field(default_factory=list)
    outlier_count: int = 0
    notes: str = ""
    missing_count: int = 0
    weakest_dimension: str = "None"
    weakest_dimension_pct: float = 100.0


@dataclass
class DatasetProfile:
    row_count: int
    column_count: int
    duplicate_rows: int
    overall_quality_score: float
    columns: list
    avg_completeness: float = 0.0
    avg_validity: float = 0.0
    avg_uniqueness: float = 0.0
    total_missing_values: int = 0
    total_outliers: int = 0
    columns_with_issues: int = 0
    worst_column: str = ""
    best_column: str = ""
    dimension_issue_counts: dict = field(default_factory=dict)


def _outliers_iqr(numeric_values: pd.Series) -> int:
    if len(numeric_values) < 4:
        return 0
    q1, q3 = numeric_values.quantile(0.25), numeric_values.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return 0
    lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    return int(((numeric_values < lower) | (numeric_values > upper)).sum())


def _clean_numeric(series: pd.Series) -> pd.Series:
    stripped = series.astype(str).str.replace(r"[^\d.\-]", "", regex=True)
    return pd.to_numeric(stripped, errors="coerce").dropna()


def _weakest_dimension(completeness_pct, validity_pct, uniqueness_pct, inferred_type):
    """Which single dimension (completeness / validity / uniqueness) is
    dragging this column's score down the most, if any."""
    dims = {"Completeness": completeness_pct, "Validity": validity_pct}
    if inferred_type in UNIQUENESS_RELEVANT_TYPES:
        dims["Uniqueness"] = uniqueness_pct
    weakest, weakest_pct = min(dims.items(), key=lambda kv: kv[1])
    if weakest_pct >= 95.0:
        return "None", 100.0
    return weakest, weakest_pct


def profile_column(series: pd.Series, column_name: str) -> ColumnProfile:
    total = len(series)
    str_series = series.astype(str)
    missing_mask = str_series.str.strip().eq("") | series.isna()
    missing = int(missing_mask.sum())
    completeness_pct = round(100 * (total - missing) / total, 1) if total else 0.0

    guess = infer_column_type(series, column_name)
    non_empty = series[~missing_mask].astype(str).str.strip()

    validity_pct = round(100 * guess.scores.get(guess.inferred_type, 0.0), 1)
    if guess.inferred_type in ("free_text", "categorical"):
        # these types are permissive by definition -> validity reflects
        # non-emptiness only, the real signal here is completeness/outliers
        validity_pct = 100.0

    invalid_samples = []
    outlier_count = 0
    notes = ""

    if guess.inferred_type == "email":
        import re
        bad = non_empty[~non_empty.str.match(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$")]
        invalid_samples = bad.head(5).tolist()
    elif guess.inferred_type == "date":
        bad_mask = non_empty.apply(lambda v: not _is_date(v))
        invalid_samples = non_empty[bad_mask].head(5).tolist()
    elif guess.inferred_type in NUMERIC_TYPES:
        numeric_vals = _clean_numeric(non_empty)
        outlier_count = _outliers_iqr(numeric_vals)
        negative_or_zero = int((numeric_vals <= 0).sum()) if guess.inferred_type in ("currency", "integer", "float") else 0
        note_parts = []
        if negative_or_zero:
            note_parts.append(f"{negative_or_zero} non-positive value(s)")
        if outlier_count:
            note_parts.append(f"{outlier_count} statistical outlier(s) (IQR method)")
        if note_parts:
            notes = "; ".join(note_parts) + " — worth a manual check."
        non_numeric = non_empty[~non_empty.index.isin(numeric_vals.index)]
        invalid_samples = non_numeric.head(5).tolist()
    elif guess.inferred_type == "categorical":
        counts = non_empty.value_counts()
        rare = counts[counts == 1]
        if len(counts) > 0 and len(rare) / len(counts) > 0.3:
            notes = f"{len(rare)} value(s) appear only once — possible free-text contamination or typos."

    uniqueness_pct = round(100 * non_empty.nunique() / len(non_empty), 1) if len(non_empty) else 0.0

    issue_count = missing + len(invalid_samples) + outlier_count
    # weighted quality score: completeness and validity matter most,
    # outliers matter less (they may be legitimate extreme values, so the
    # penalty is capped) — a rare-category note on a categorical column
    # also shaves a little off since it signals possible free-text noise.
    outlier_penalty = min(15.0, outlier_count * 1.5)
    quality_score = round(
        max(
            0.0,
            0.5 * completeness_pct
            + 0.4 * validity_pct
            + 0.1 * min(uniqueness_pct, 100)
            - outlier_penalty
            - (3 if notes else 0),
        ),
        1,
    )
    quality_score = min(quality_score, 100.0)

    weakest_dimension, weakest_dimension_pct = _weakest_dimension(
        completeness_pct, validity_pct, uniqueness_pct, guess.inferred_type
    )

    return ColumnProfile(
        column=column_name,
        inferred_type=guess.inferred_type,
        confidence=round(guess.confidence, 2),
        completeness_pct=completeness_pct,
        validity_pct=validity_pct,
        uniqueness_pct=uniqueness_pct,
        issue_count=issue_count,
        quality_score=quality_score,
        sample_invalid=[str(x) for x in invalid_samples],
        outlier_count=outlier_count,
        notes=notes,
        missing_count=missing,
        weakest_dimension=weakest_dimension,
        weakest_dimension_pct=weakest_dimension_pct,
    )


def profile_dataset(df: pd.DataFrame) -> DatasetProfile:
    duplicate_rows = int(df.duplicated(keep="first").sum())
    columns = [profile_column(df[col], col) for col in df.columns]
    overall = round(sum(c.quality_score for c in columns) / len(columns), 1) if columns else 0.0
    # duplicate rows drag the overall score down a bit, dataset-wide
    dup_penalty = min(15.0, round(100 * duplicate_rows / len(df), 1)) if len(df) else 0.0
    overall = round(max(0.0, overall - dup_penalty * 0.3), 1)

    n = len(columns)
    avg_completeness = round(sum(c.completeness_pct for c in columns) / n, 1) if n else 0.0
    avg_validity = round(sum(c.validity_pct for c in columns) / n, 1) if n else 0.0
    avg_uniqueness = round(sum(c.uniqueness_pct for c in columns) / n, 1) if n else 0.0
    total_missing_values = sum(c.missing_count for c in columns)
    total_outliers = sum(c.outlier_count for c in columns)
    columns_with_issues = sum(1 for c in columns if c.quality_score < 85)
    worst_column = min(columns, key=lambda c: c.quality_score).column if columns else ""
    best_column = max(columns, key=lambda c: c.quality_score).column if columns else ""
    dimension_issue_counts = dict(
        Counter(c.weakest_dimension for c in columns if c.weakest_dimension != "None")
    )

    return DatasetProfile(
        row_count=len(df),
        column_count=len(df.columns),
        duplicate_rows=duplicate_rows,
        overall_quality_score=overall,
        columns=columns,
        avg_completeness=avg_completeness,
        avg_validity=avg_validity,
        avg_uniqueness=avg_uniqueness,
        total_missing_values=total_missing_values,
        total_outliers=total_outliers,
        columns_with_issues=columns_with_issues,
        worst_column=worst_column,
        best_column=best_column,
        dimension_issue_counts=dimension_issue_counts,
    )
