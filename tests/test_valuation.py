"""Tests for implied valuation — target metrics × peer multiples."""
import pandas as pd
import pytest
from valuation import compute_implied, summary_range


def _summary():
    # Minimal summary_stats-shaped frame
    return pd.DataFrame([
        {"metric": "ev_sales",  "q1": 2.0, "median": 3.0, "q3": 4.0},
        {"metric": "ev_ebitda", "q1": 8.0, "median": 10.0, "q3": 12.0},
        {"metric": "ev_ebit",   "q1": 12.0, "median": 15.0, "q3": 18.0},
        {"metric": "pe",        "q1": 14.0, "median": 18.0, "q3": 22.0},
    ])


def test_implied_ev_from_ebitda_median():
    target = {"revenue": 1000, "ebitda": 200, "ebit": 150, "net_income": 90, "net_debt": 100}
    rows = compute_implied(target, _summary())
    ev_eb = next(r for r in rows if r["multiple"] == "ev_ebitda")
    # base EV = EBITDA · median EV/EBITDA = 200 · 10 = 2000
    assert ev_eb["base_ev"] == pytest.approx(2000)


def test_implied_equity_is_ev_minus_net_debt():
    target = {"revenue": 1000, "ebitda": 200, "ebit": 150, "net_income": 90, "net_debt": 100}
    rows = compute_implied(target, _summary())
    ev_eb = next(r for r in rows if r["multiple"] == "ev_ebitda")
    assert ev_eb["base_equity"] == pytest.approx(ev_eb["base_ev"] - 100)


def test_pe_bridges_equity_to_ev():
    # P/E gives equity directly; EV = equity + net debt
    target = {"revenue": 1000, "ebitda": 200, "ebit": 150, "net_income": 90, "net_debt": 100}
    rows = compute_implied(target, _summary())
    pe = next(r for r in rows if r["multiple"] == "pe")
    # base equity = NI · median P/E = 90 · 18 = 1620
    assert pe["base_equity"] == pytest.approx(1620)
    assert pe["base_ev"] == pytest.approx(1620 + 100)


def test_low_base_high_ordering():
    target = {"revenue": 1000, "ebitda": 200, "ebit": 150, "net_income": 90, "net_debt": 100}
    rows = compute_implied(target, _summary())
    for r in rows:
        assert r["low_ev"] <= r["base_ev"] <= r["high_ev"]


def test_missing_operand_skips_multiple():
    # No EBITDA → no EV/EBITDA row
    target = {"revenue": 1000, "ebitda": None, "ebit": 150, "net_income": 90, "net_debt": 100}
    rows = compute_implied(target, _summary())
    assert all(r["multiple"] != "ev_ebitda" for r in rows)


def test_summary_range_spans_all_methods():
    target = {"revenue": 1000, "ebitda": 200, "ebit": 150, "net_income": 90, "net_debt": 100}
    rows = compute_implied(target, _summary())
    rng = summary_range(rows)
    evs = [v for r in rows for v in (r["low_ev"], r["base_ev"], r["high_ev"]) if v]
    assert rng["ev_min"] == pytest.approx(min(evs))
    assert rng["ev_max"] == pytest.approx(max(evs))
