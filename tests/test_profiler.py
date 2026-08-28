import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

from src.profiler import profile_dataset


def test_profile_dataset_basic_shape():
    df = pd.DataFrame({
        "id": ["1", "2", "3", "4"],
        "email": ["a@b.com", "bad-email", "c@d.com", "e@f.com"],
        "status": ["Active", "Active", "On Hold", "Active"],
    })
    profile = profile_dataset(df)
    assert profile.row_count == 4
    assert profile.column_count == 3
    assert len(profile.columns) == 3
    email_col = next(c for c in profile.columns if c.column == "email")
    assert email_col.inferred_type == "email"
    assert email_col.issue_count >= 1


def test_duplicate_rows_detected():
    df = pd.DataFrame({
        "id": ["1", "2", "2"],
        "name": ["A", "B", "B"],
    })
    profile = profile_dataset(df)
    assert profile.duplicate_rows == 1


def test_missing_values_reduce_completeness():
    df = pd.DataFrame({"col": ["1", "", "3", ""]})
    profile = profile_dataset(df)
    col = profile.columns[0]
    assert col.completeness_pct == 50.0


def test_perfect_column_scores_high():
    df = pd.DataFrame({"email": ["a@b.com", "c@d.com", "e@f.com"]})
    profile = profile_dataset(df)
    col = profile.columns[0]
    assert col.quality_score >= 90


def test_overall_score_in_valid_range():
    df = pd.DataFrame({
        "a": ["1", "2", "", "4"],
        "b": ["x@y.com", "bad", "z@w.com", "q@w.com"],
    })
    profile = profile_dataset(df)
    assert 0 <= profile.overall_quality_score <= 100
