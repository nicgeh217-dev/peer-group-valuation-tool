"""Implied Valuation: target metrics * peer multiples = implied value ranges.

Output for each available multiple:
- Low case   = target operand * peer Q1
- Base case  = target operand * peer Median
- High case  = target operand * peer Q3

EV-based multiples (EV/Sales, EV/EBITDA, EV/EBIT) → Implied EV, then bridged
to Equity by subtracting Net Debt.
P/E → Implied Equity directly; EV derived by adding Net Debt back.
"""
from __future__ import annotations

import pandas as pd

# multiple → (target operand field, output type)
MULTIPLE_INPUTS = {
    "ev_sales":  ("revenue",    "EV"),
    "ev_ebitda": ("ebitda",     "EV"),
    "ev_ebit":   ("ebit",       "EV"),
    "pe":        ("net_income", "Equity"),
}

# Pretty labels for the Excel tab
MULTIPLE_LABELS = {
    "ev_sales":  "EV / Sales",
    "ev_ebitda": "EV / EBITDA",
    "ev_ebit":   "EV / EBIT",
    "pe":        "P / E",
}


def compute_implied(target: dict, summary_df: pd.DataFrame) -> list[dict]:
    """target: {revenue, ebitda, ebit, net_income, net_debt} in target currency.
    summary_df: from multiples.summary_stats (columns: metric, q1, median, q3).
    Returns list of dicts per multiple.
    """
    summary_lookup = {row["metric"]: row for _, row in summary_df.iterrows()}
    net_debt = target.get("net_debt") or 0
    results = []

    for mult, (operand, kind) in MULTIPLE_INPUTS.items():
        target_val = target.get(operand)
        stats = summary_lookup.get(mult)
        if target_val is None or target_val == 0 or stats is None:
            continue
        median = stats.get("median")
        q1 = stats.get("q1")
        q3 = stats.get("q3")
        if median is None or pd.isna(median):
            continue

        low_val  = (target_val * q1)     if q1     and not pd.isna(q1)     else None
        base_val = (target_val * median)
        high_val = (target_val * q3)     if q3     and not pd.isna(q3)     else None

        if kind == "EV":
            low_ev,  base_ev,  high_ev  = low_val,  base_val,  high_val
            low_eq   = (low_ev  - net_debt) if low_ev  is not None else None
            base_eq  = base_ev  - net_debt
            high_eq  = (high_ev - net_debt) if high_ev is not None else None
        else:  # P/E → equity first, bridge up to EV
            low_eq,  base_eq,  high_eq  = low_val,  base_val,  high_val
            low_ev   = (low_eq  + net_debt) if low_eq  is not None else None
            base_ev  = base_eq  + net_debt
            high_ev  = (high_eq + net_debt) if high_eq is not None else None

        results.append({
            "multiple":     mult,
            "label":        MULTIPLE_LABELS[mult],
            "operand":      operand,
            "operand_val":  target_val,
            "q1":           q1,
            "median":       median,
            "q3":           q3,
            "low_ev":       low_ev,
            "base_ev":      base_ev,
            "high_ev":      high_ev,
            "low_equity":   low_eq,
            "base_equity":  base_eq,
            "high_equity":  high_eq,
        })

    return results


def summary_range(results: list[dict]) -> dict:
    """Returns the overall min/max EV and Equity across all multiples (Football-Field range)."""
    if not results:
        return {}
    evs = [v for r in results for v in (r["low_ev"], r["base_ev"], r["high_ev"]) if v is not None]
    eqs = [v for r in results for v in (r["low_equity"], r["base_equity"], r["high_equity"]) if v is not None]
    return {
        "ev_min":      min(evs) if evs else None,
        "ev_max":      max(evs) if evs else None,
        "equity_min":  min(eqs) if eqs else None,
        "equity_max":  max(eqs) if eqs else None,
        "ev_median":   sorted([r["base_ev"]     for r in results if r["base_ev"]     is not None])[len(results)//2] if results else None,
        "equity_median": sorted([r["base_equity"] for r in results if r["base_equity"] is not None])[len(results)//2] if results else None,
    }
