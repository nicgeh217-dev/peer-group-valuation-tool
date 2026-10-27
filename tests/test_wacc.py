"""Tests for WACC / CAPM logic — the core of the DCF discount rate."""
import math
import pytest
from wacc import (
    unlever_beta, relever_beta, peer_unlevered_betas, compute_wacc,
)


# ---------- Beta unlevering / relevering (Hamada) ----------

def test_unlever_beta_no_debt():
    # With zero debt, unlevered beta == levered beta
    assert unlever_beta(1.2, tax=0.25, de_ratio=0.0) == pytest.approx(1.2)


def test_unlever_beta_with_debt():
    # βu = βl / (1 + (1-tax)·D/E) = 1.5 / (1 + 0.75·0.5) = 1.5 / 1.375
    assert unlever_beta(1.5, tax=0.25, de_ratio=0.5) == pytest.approx(1.5 / 1.375)


def test_relever_is_inverse_of_unlever():
    # Unlever then relever at the same D/E must return the original beta
    bl, tax, de = 1.4, 0.25, 0.6
    bu = unlever_beta(bl, tax, de)
    assert relever_beta(bu, tax, de) == pytest.approx(bl)


def test_higher_debt_lowers_unlevered_beta():
    # A distressed (high D/E) peer must produce a LOWER unlevered beta
    low_de = unlever_beta(1.5, 0.25, 0.3)
    high_de = unlever_beta(1.5, 0.25, 3.0)
    assert high_de < low_de


# ---------- Peer beta table + D/E outlier flag ----------

def _peer(sym, beta, mcap, nd):
    return {"symbol": sym, "name": sym, "beta": beta, "market_cap": mcap, "net_debt": nd}


def test_peer_betas_skip_missing():
    peers = [_peer("A", None, 1e9, 1e8), _peer("B", 1.0, None, 1e8), _peer("C", 1.2, 5e9, 1e9)]
    out = peer_unlevered_betas(peers, tax=0.25)
    assert [p["symbol"] for p in out] == ["C"]  # A has no beta, B no market cap


def test_peer_betas_dedup_dual_class():
    # Same company name twice (e.g. Sartorius ordinary + pref) counted once
    peers = [_peer("SRT.DE", 1.4, 1e10, 3e9), _peer("SRT3.DE", 1.4, 1.2e10, 3e9)]
    peers[0]["name"] = peers[1]["name"] = "Sartorius AG"
    out = peer_unlevered_betas(peers, tax=0.25)
    assert len(out) == 1


def test_de_outlier_flagged():
    # D/E above 200% must be flagged
    peers = [_peer("X", 1.6, 0.6e9, 2.2e9)]  # net_debt 2.2B / mcap 0.6B = 366% D/E
    w = compute_wacc(peers + [_peer("Y", 1.0, 5e9, 1e9), _peer("Z", 1.1, 4e9, 0.5e9)],
                     rf=0.025, mrp=0.055, tax=0.25, kd=0.05, target_de=0.5)
    flagged = [p for p in w["peer_betas"] if p.get("de_outlier")]
    assert any(p["symbol"] == "X" for p in flagged)
    assert w["n_de_outliers"] >= 1


# ---------- Full WACC ----------

def test_wacc_formula_end_to_end():
    # Single clean peer, beta 1.0, no debt → unlevered = relevered to target D/E
    peers = [_peer("A", 1.0, 1e10, 0.0), _peer("B", 1.0, 1e10, 0.0)]
    w = compute_wacc(peers, rf=0.03, mrp=0.05, tax=0.25, kd=0.06, target_de=0.0)
    # No debt → WACC == cost of equity == Rf + beta·MRP = 0.03 + 1.0·0.05 = 0.08
    assert w["cost_of_equity"] == pytest.approx(0.08)
    assert w["wacc"] == pytest.approx(0.08)


def test_wacc_weights_sum_to_one():
    peers = [_peer("A", 1.1, 5e9, 1e9), _peer("B", 0.9, 6e9, 2e9)]
    w = compute_wacc(peers, rf=0.025, mrp=0.055, tax=0.25, kd=0.05, target_de=0.5)
    assert w["weight_equity"] + w["weight_debt"] == pytest.approx(1.0)


def test_size_premium_raises_cost_of_equity():
    peers = [_peer("A", 1.0, 1e10, 0.0)]
    base = compute_wacc(peers, 0.025, 0.055, 0.25, 0.05, 0.3, size_premium=0.0)
    prem = compute_wacc(peers, 0.025, 0.055, 0.25, 0.05, 0.3, size_premium=0.02)
    assert prem["cost_of_equity"] == pytest.approx(base["cost_of_equity"] + 0.02)
    assert prem["wacc"] > base["wacc"]


def test_wacc_empty_peers_returns_empty():
    assert compute_wacc([], 0.025, 0.055, 0.25, 0.05, 0.5) == {}
