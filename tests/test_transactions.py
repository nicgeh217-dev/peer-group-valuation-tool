"""Tests for precedent transactions — deal multiples, sector filter, implied value."""
import pytest
from transactions import (
    Deal, deal_multiple_stats, implied_from_transactions, filter_by_sector, _quantile,
)


def _deal(target, ev, rev, ebitda, ebit, sector=None):
    return Deal(date="2021", target=target, acquirer="Buyer",
                ev=ev, revenue=rev, ebitda=ebitda, ebit=ebit, sector=sector)


# ---------- Deal multiple computation ----------

def test_deal_multiple_properties():
    d = _deal("X", ev=1000, rev=500, ebitda=100, ebit=80)
    assert d.ev_sales == pytest.approx(2.0)
    assert d.ev_ebitda == pytest.approx(10.0)
    assert d.ev_ebit == pytest.approx(12.5)


def test_deal_multiple_none_when_missing():
    d = _deal("X", ev=1000, rev=None, ebitda=0, ebit=80)
    assert d.ev_sales is None        # no revenue
    assert d.ev_ebitda is None       # division by zero guarded


def test_deal_stats_median():
    deals = [
        _deal("A", 1000, 500, 100, 80),   # EV/EBITDA 10
        _deal("B", 1500, 500, 100, 80),   # EV/EBITDA 15
        _deal("C", 2000, 500, 100, 80),   # EV/EBITDA 20
    ]
    stats = deal_multiple_stats(deals)
    assert stats["ev_ebitda"]["median"] == pytest.approx(15.0)
    assert stats["ev_ebitda"]["n"] == 3


def test_deal_stats_drops_nonpositive():
    deals = [_deal("A", 1000, 500, 100, 80), _deal("B", -500, 500, 100, 80)]
    stats = deal_multiple_stats(deals)
    # Negative EV/EBITDA dropped → only one valid value
    assert stats["ev_ebitda"]["n"] == 1


# ---------- Sector filter ----------

def test_sector_filter_keyword_match():
    deals = [
        _deal("A", 1, 1, 1, 1, sector="Payments"),
        _deal("B", 1, 1, 1, 1, sector="Auto Parts"),
    ]
    out = filter_by_sector(deals, "Payments / Fintech — Global")
    assert [d.target for d in out] == ["A"]


def test_sector_filter_keeps_untagged():
    # Deals without a sector tag are treated as universal
    deals = [_deal("A", 1, 1, 1, 1, sector=None)]
    out = filter_by_sector(deals, "Payments / Fintech — Global")
    assert len(out) == 1


def test_sector_filter_no_universe_returns_all():
    deals = [_deal("A", 1, 1, 1, 1, sector="Payments")]
    assert len(filter_by_sector(deals, None)) == 1


# ---------- Implied valuation from deals ----------

def test_implied_from_transactions():
    deals = [_deal("A", 1000, 500, 100, 80), _deal("B", 1500, 500, 100, 80)]
    stats = deal_multiple_stats(deals)
    target = {"revenue": 1000, "ebitda": 200, "ebit": 150, "net_debt": 50}
    rows = implied_from_transactions(target, stats)
    ev_eb = next(r for r in rows if r["multiple"] == "ev_ebitda")
    # median EV/EBITDA = 12.5; base EV = 200 · 12.5 = 2500
    assert ev_eb["base_ev"] == pytest.approx(2500)
    assert ev_eb["base_equity"] == pytest.approx(2500 - 50)


# ---------- Quantile helper ----------

def test_quantile_interpolates():
    vals = [1.0, 2.0, 3.0, 4.0]
    assert _quantile(vals, 0.5) == pytest.approx(2.5)
    assert _quantile(vals, 0.25) == pytest.approx(1.75)
