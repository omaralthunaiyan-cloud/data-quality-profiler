import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

from src.type_inference import infer_column_type


def test_detects_email_column_with_some_noise():
    s = pd.Series(["a@b.com", "c@d.com", "not-an-email", "e@f.com"])
    g = infer_column_type(s, "email")
    assert g.inferred_type == "email"
    assert 0.6 <= g.confidence <= 1.0


def test_detects_date_column_mixed_formats():
    s = pd.Series(["2024-01-05", "05/02/2024", "2024-03-10"])
    g = infer_column_type(s, "hire_date")
    assert g.inferred_type == "date"


def test_dates_not_misclassified_as_phone():
    s = pd.Series(["2024-01-05", "2024-02-06", "2024-03-07", "2024-04-08"])
    g = infer_column_type(s, "start_date")
    assert g.inferred_type == "date"


def test_detects_categorical_column():
    s = pd.Series(["Active", "Active", "On Hold", "Active", "Completed", "Active"])
    g = infer_column_type(s, "status")
    assert g.inferred_type == "categorical"


def test_detects_integer_column():
    s = pd.Series(["1", "2", "3", "42", "100"])
    g = infer_column_type(s, "count")
    assert g.inferred_type == "integer"


def test_falls_back_to_free_text_for_unstructured_names():
    s = pd.Series(["Omar Althunaiyan", "Sara Al-Qahtani", "Faisal Al-Otaibi"])
    g = infer_column_type(s, "full_name")
    assert g.inferred_type in ("free_text", "categorical")


def test_empty_series_does_not_crash():
    s = pd.Series(["", "", ""])
    g = infer_column_type(s, "blank")
    assert g.inferred_type is not None
