"""Tests for the DCF engine — projection, terminal value, sensitivity."""
import pytest
from wacc import dcf, dcf_sensitivity, fcf_bridge


def test_dcf_basic_structure():
    d = dcf(base_fcf=100.0, growth=0.05, terminal_growth=0.02, wacc=0.10, years=5)
    assert len(d["fcfs"]) == 5
    assert len(d["pv_fcfs"]) == 5
    assert d["enterprise_value"] > 0


def test_dcf_first_year_fcf_grows():
    # Year 1 FCF = base · (1+growth)
    d = dcf(base_fcf=100.0, growth=0.05, terminal_growth=0.02, wacc=0.10, years=5)
    assert d["fcfs"][0] == pytest.approx(105.0)


def test_dcf_terminal_value_gordon_growth():
    # TV = FCF_N·(1+tg) / (wacc - tg)
    d = dcf(base_fcf=100.0, growth=0.0, terminal_growth=0.02, wacc=0.10, years=5)
    fcf5 = d["fcfs"][-1]  # growth 0 → all years = 100
    expected_tv = fcf5 * 1.02 / (0.10 - 0.02)
    assert d["terminal_value"] == pytest.approx(expected_tv)


def test_dcf_pv_discounting():
    # PV of year-1 FCF = FCF1 / (1+wacc)^1
    d = dcf(base_fcf=100.0, growth=0.0, terminal_growth=0.02, wacc=0.10, years=5)
    assert d["pv_fcfs"][0] == pytest.approx(100.0 / 1.10)


def test_dcf_lower_wacc_gives_higher_ev():
    low = dcf(100.0, 0.04, 0.02, wacc=0.08, years=5)
    high = dcf(100.0, 0.04, 0.02, wacc=0.12, years=5)
    assert low["enterprise_value"] > high["enterprise_value"]


def test_dcf_rejects_wacc_below_terminal_growth():
    # WACC must exceed terminal growth or Gordon Growth blows up
    d = dcf(100.0, 0.04, terminal_growth=0.05, wacc=0.04, years=5)
    assert d["enterprise_value"] is None
    assert "error" in d


def test_dcf_tv_share_reported():
    d = dcf(100.0, 0.04, 0.02, 0.09, years=5)
    assert 0.0 < d["tv_pct_of_ev"] < 1.0


def test_dcf_sensitivity_grid_shape():
    grid = dcf_sensitivity(100.0, 0.04, 0.02, base_wacc=0.09, years=5)
    assert len(grid) == 9  # 3 WACC × 3 terminal-growth
    # All entries carry a wacc, terminal_growth, and an EV
    for cell in grid:
        assert "wacc" in cell and "terminal_growth" in cell


# ---- FCF bridge (FCFF from operating metrics) ----

def test_fcf_bridge_formula():
    # FCF = EBIT·(1−tax) + D&A − Capex − ΔNWC
    # 500·(1−0.25) + 400 − 240 − 60 = 375 + 400 − 240 − 60 = 475
    assert fcf_bridge(ebit=500, tax=0.25, da=400, capex=240, delta_nwc=60) == pytest.approx(475.0)


def test_fcf_bridge_da_from_ebitda_minus_ebit():
    # D&A is EBITDA − EBIT in practice: EBITDA 900, EBIT 500 → D&A 400
    ebitda, ebit = 900, 500
    da = ebitda - ebit
    fcf = fcf_bridge(ebit=ebit, tax=0.25, da=da, capex=0, delta_nwc=0)
    assert fcf == pytest.approx(ebit * 0.75 + 400)


def test_fcf_bridge_percent_of_revenue_inputs():
    # Capex as % of revenue, ΔNWC as % of the revenue increment (revenue·growth)
    revenue, growth = 6000.0, 0.04
    capex_pct, nwc_pct = 0.05, 0.10
    capex = revenue * capex_pct                 # 300
    delta_nwc = revenue * growth * nwc_pct       # 6000·0.04·0.10 = 24
    fcf = fcf_bridge(ebit=500, tax=0.25, da=400, capex=capex, delta_nwc=delta_nwc)
    assert fcf == pytest.approx(500 * 0.75 + 400 - 300 - 24)  # 451


def test_dcf_echoes_bridge_for_export():
    bridge = {"ebit": 500, "tax": 0.25, "da": 400, "capex": 300, "delta_nwc": 24}
    base = fcf_bridge(500, 0.25, 400, 300, 24)
    d = dcf(base, 0.04, 0.02, 0.09, years=5, bridge=bridge)
    assert d["bridge"] == bridge
    assert d["base_fcf"] == pytest.approx(base)


def test_dcf_bridge_none_by_default():
    d = dcf(100.0, 0.04, 0.02, 0.09, years=5)
    assert d["bridge"] is None
