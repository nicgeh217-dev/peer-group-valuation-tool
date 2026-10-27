"""Tests for outlier detection (Tukey IQR + negatives) and summary statistics."""
import pandas as pd
import pytest
from multiples import flag_outliers, summary_stats


def _df(ev_ebitda_values):
    return pd.DataFrame({
        "symbol": [f"P{i}" for i in range(len(ev_ebitda_values))],
        "ev_sales": [3.0] * len(ev_ebitda_values),
        "ev_ebitda": ev_ebitda_values,
        "ev_ebit": [15.0] * len(ev_ebitda_values),
        "pe": [18.0] * len(ev_ebitda_values),
        "ebitda_margin": [0.2] * len(ev_ebitda_values),
        "ebit_margin": [0.15] * len(ev_ebitda_values),
    })


def test_negative_multiple_flagged():
    df = _df([10.0, 12.0, -5.0, 11.0])
    out = flag_outliers(df)
    # The negative EV/EBITDA row must carry a flag
    assert "ev_ebitda<0" in out.iloc[2]["outlier_flags"]
    assert out.iloc[2]["is_outlier"]


def test_iqr_outlier_flagged():
    # One extreme value among tight cluster → IQR outlier
    df = _df([10.0, 11.0, 12.0, 13.0, 200.0])
    out = flag_outliers(df)
    assert out.iloc[4]["is_outlier"]            # 200x is the outlier
    assert not out.iloc[0]["is_outlier"]        # 10x is normal


def test_clean_values_not_flagged():
    df = _df([9.0, 10.0, 11.0, 12.0])
    out = flag_outliers(df)
    assert not out["is_outlier"].any()


def test_summary_excludes_outliers():
    df = _df([10.0, 11.0, 12.0, 13.0, 200.0])
    df = flag_outliers(df)
    df["included"] = True
    summ = summary_stats(df, exclude_outliers=True)
    eb = summ[summ["metric"] == "ev_ebitda"].iloc[0]
    # Median of the clean 4 (10,11,12,13) = 11.5, NOT pulled up by 200
    assert eb["median"] == pytest.approx(11.5)
    assert eb["n"] == 4


def test_summary_drops_negative_multiples():
    df = _df([10.0, 12.0, -5.0, 11.0])
    df = flag_outliers(df)
    df["included"] = True
    summ = summary_stats(df, exclude_outliers=False)
    eb = summ[summ["metric"] == "ev_ebitda"].iloc[0]
    # Negative dropped from multiple stats → median of 10,11,12
    assert eb["median"] == pytest.approx(11.0)
