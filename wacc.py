"""WACC via peer betas (CAPM) + a light DCF.

Flow:
  1. Peer levered betas (Yahoo) → unlever per peer:  βu = βl / (1 + (1-Tax)·D/E)
  2. Median unlevered beta of the peer group
  3. Relever to target capital structure:  βl_target = βu · (1 + (1-Tax)·D/E_target)
  4. CAPM cost of equity:  Ke = Rf + βl_target · MRP
  5. WACC = E/V·Ke + D/V·Kd·(1-Tax)

DCF (light):
  - Project FCF over N years from a base FCF + growth rate
  - Discount at WACC
  - Terminal value via Gordon Growth:  TV = FCF_N·(1+g) / (WACC - g)
  - Enterprise Value = PV(FCFs) + PV(TV)
"""
from __future__ import annotations

import statistics


def unlever_beta(beta_l: float, tax: float, de_ratio: float) -> float:
    return beta_l / (1 + (1 - tax) * de_ratio)


def relever_beta(beta_u: float, tax: float, de_ratio: float) -> float:
    return beta_u * (1 + (1 - tax) * de_ratio)


def peer_unlevered_betas(peers: list[dict], tax: float) -> list[dict]:
    """peers: dicts with symbol, beta, net_debt, market_cap.
    Returns per-peer dict with levered/unlevered beta and D/E.
    Skips peers without a positive beta or market cap.
    Dedups dual-class listings (same company name) to avoid double-counting in median.
    """
    out = []
    seen_names = set()
    for p in peers:
        beta_l = p.get("beta")
        mcap = p.get("market_cap")
        net_debt = p.get("net_debt")
        if beta_l is None or mcap is None or mcap <= 0:
            continue
        # Dedup dual-class shares (e.g. Sartorius Stamm + Vz) by company name
        name_key = (p.get("name") or p.get("symbol") or "").strip().lower()
        if name_key in seen_names:
            continue
        seen_names.add(name_key)
        # Net debt can be negative (net cash). Floor D at 0 for D/E.
        debt = max(net_debt or 0, 0)
        de = debt / mcap if mcap else 0
        beta_u = unlever_beta(beta_l, tax, de)
        out.append({
            "symbol": p.get("symbol"),
            "name": p.get("name"),
            "beta_levered": beta_l,
            "de_ratio": de,
            "beta_unlevered": beta_u,
        })
    return out


DE_OUTLIER_THRESHOLD = 2.0   # D/E above 200% is treated as distorted (distressed cap structure)


def compute_wacc(
    peers: list[dict],
    rf: float,
    mrp: float,
    tax: float,
    kd: float,
    target_de: float,
    size_premium: float = 0.0,
) -> dict:
    """Returns full WACC breakdown. All rates as decimals (0.025 = 2.5%).
    target_de: target Debt/Equity ratio.
    size_premium: optional add-on to cost of equity (decimal, e.g. 0.02 = +2%).
    """
    peer_betas = peer_unlevered_betas(peers, tax)
    if not peer_betas:
        return {}

    # Flag D/E outliers (distorted unlevering)
    for p in peer_betas:
        p["de_outlier"] = p["de_ratio"] > DE_OUTLIER_THRESHOLD

    unlevered_list = [p["beta_unlevered"] for p in peer_betas]
    median_bu = statistics.median(unlevered_list)
    mean_bu = statistics.mean(unlevered_list)

    # Diagnostic: median EXCLUDING D/E outliers (so analyst sees the distortion effect)
    clean_list = [p["beta_unlevered"] for p in peer_betas if not p["de_outlier"]]
    median_bu_clean = statistics.median(clean_list) if len(clean_list) >= 2 else None
    n_outliers = sum(1 for p in peer_betas if p["de_outlier"])

    # Relever to target structure
    beta_target = relever_beta(median_bu, tax, target_de)

    # CAPM (+ optional size premium)
    ke = rf + beta_target * mrp + size_premium

    # Capital structure weights from D/E
    # D/E = target_de → E weight = 1/(1+D/E), D weight = D/E/(1+D/E)
    w_e = 1 / (1 + target_de)
    w_d = target_de / (1 + target_de)
    after_tax_kd = kd * (1 - tax)
    wacc = w_e * ke + w_d * after_tax_kd

    # What WACC would be if we used the D/E-outlier-cleaned beta (diagnostic)
    wacc_clean = None
    if median_bu_clean is not None and n_outliers > 0:
        beta_t_clean = relever_beta(median_bu_clean, tax, target_de)
        ke_clean = rf + beta_t_clean * mrp + size_premium
        wacc_clean = w_e * ke_clean + w_d * after_tax_kd

    return {
        "peer_betas": peer_betas,
        "median_unlevered_beta": median_bu,
        "mean_unlevered_beta": mean_bu,
        "median_unlevered_beta_clean": median_bu_clean,
        "n_de_outliers": n_outliers,
        "de_outlier_threshold": DE_OUTLIER_THRESHOLD,
        "wacc_clean": wacc_clean,
        "target_de": target_de,
        "relevered_beta": beta_target,
        "rf": rf,
        "mrp": mrp,
        "tax": tax,
        "kd": kd,
        "size_premium": size_premium,
        "after_tax_kd": after_tax_kd,
        "cost_of_equity": ke,
        "weight_equity": w_e,
        "weight_debt": w_d,
        "wacc": wacc,
    }


def fcf_bridge(ebit: float, tax: float, da: float,
               capex: float, delta_nwc: float) -> float:
    """Unlevered free cash flow (FCFF) from operating metrics:

        FCF = EBIT·(1−tax) + D&A − Capex − ΔNWC
              └─ NOPAT ───┘   └ non-cash ┘  └ reinvestment ┘

    All amounts in currency units, tax as a decimal. This is the standard bridge a
    DCF starts from — it derives the free cash flow instead of taking it as a raw
    input, so the analyst can defend where the cash flow comes from.
    """
    return ebit * (1 - tax) + da - capex - delta_nwc


def dcf(
    base_fcf: float,
    growth: float,
    terminal_growth: float,
    wacc: float,
    years: int = 5,
    bridge: dict | None = None,
) -> dict:
    """Light DCF. base_fcf in currency units, growth/terminal_growth/wacc as decimals.
    Returns enterprise value + components.

    bridge: optional FCFF-bridge inputs (ebit, tax, da, capex, delta_nwc, revenue,
    growth, capex_pct, nwc_pct). Echoed back so the Excel export can rebuild the
    cash-flow bridge as live formulas above the projection. Does not change the math:
    the caller passes base_fcf = fcf_bridge(...) and the same value here.
    """
    if wacc is None or wacc <= terminal_growth:
        return {"error": "WACC must exceed terminal growth", "enterprise_value": None}

    fcfs = []
    pv_fcfs = []
    fcf = base_fcf
    for yr in range(1, years + 1):
        fcf = fcf * (1 + growth)
        pv = fcf / ((1 + wacc) ** yr)
        fcfs.append(fcf)
        pv_fcfs.append(pv)

    # Terminal value at year N (Gordon Growth), discounted back
    tv = fcfs[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_tv = tv / ((1 + wacc) ** years)

    ev = sum(pv_fcfs) + pv_tv
    return {
        "years": years,
        # raw inputs echoed back so the Excel export can rebuild the model with
        # live formulas that reference an assumptions block instead of dead numbers
        "base_fcf": base_fcf,
        "growth": growth,
        "terminal_growth": terminal_growth,
        "wacc": wacc,
        "bridge": bridge,
        "fcfs": fcfs,
        "pv_fcfs": pv_fcfs,
        "terminal_value": tv,
        "pv_terminal_value": pv_tv,
        "sum_pv_fcf": sum(pv_fcfs),
        "enterprise_value": ev,
        "tv_pct_of_ev": pv_tv / ev if ev else None,
    }


def dcf_sensitivity(
    base_fcf: float,
    growth: float,
    terminal_growth: float,
    base_wacc: float,
    years: int = 5,
    wacc_delta: float = 0.01,
    tg_delta: float = 0.005,
) -> list[dict]:
    """Returns a small WACC × terminal-growth sensitivity grid of EVs."""
    grid = []
    for w in [base_wacc - wacc_delta, base_wacc, base_wacc + wacc_delta]:
        for tg in [terminal_growth - tg_delta, terminal_growth, terminal_growth + tg_delta]:
            res = dcf(base_fcf, growth, tg, w, years)
            grid.append({
                "wacc": w,
                "terminal_growth": tg,
                "enterprise_value": res.get("enterprise_value"),
            })
    return grid
