"""Screening logic: filter candidates by size, margin, sector proximity."""
from __future__ import annotations

import pandas as pd


def apply_filters(
    df: pd.DataFrame,
    revenue_min: float | None,
    revenue_max: float | None,
    ebitda_margin_min: float | None,
    ebitda_margin_max: float | None,
) -> pd.DataFrame:
    """Returns df with `excluded_reason` column. Does not drop rows by default —
    excluded rows are kept but marked, so the analyst sees what was filtered.
    """
    df = df.copy()
    reasons = [[] for _ in range(len(df))]

    rev = pd.to_numeric(df.get("revenue"), errors="coerce")
    margin = pd.to_numeric(df.get("ebitda_margin"), errors="coerce")

    for i in range(len(df)):
        r = rev.iloc[i] if i < len(rev) else None
        m = margin.iloc[i] if i < len(margin) else None
        if revenue_min is not None and (pd.isna(r) or r < revenue_min):
            reasons[i].append(f"rev<{revenue_min:.0f}")
        if revenue_max is not None and (pd.isna(r) or r > revenue_max):
            reasons[i].append(f"rev>{revenue_max:.0f}")
        if ebitda_margin_min is not None and (pd.isna(m) or m < ebitda_margin_min):
            reasons[i].append(f"margin<{ebitda_margin_min:.0%}")
        if ebitda_margin_max is not None and (pd.isna(m) or m > ebitda_margin_max):
            reasons[i].append(f"margin>{ebitda_margin_max:.0%}")

    df["excluded_reason"] = ["; ".join(r) if r else "" for r in reasons]
    df["included"] = df["excluded_reason"] == ""
    return df


def rank_by_similarity(df: pd.DataFrame, target_revenue: float | None) -> pd.DataFrame:
    """Sort by absolute log-distance of revenue from target. Falls back to revenue desc."""
    df = df.copy()
    rev = pd.to_numeric(df.get("revenue"), errors="coerce")

    if target_revenue and target_revenue > 0:
        import numpy as np
        log_target = np.log(target_revenue)
        distance = (np.log(rev.where(rev > 0)) - log_target).abs()
        df["_distance"] = distance
        df = df.sort_values(by=["included", "_distance"], ascending=[False, True], na_position="last")
        df = df.drop(columns=["_distance"])
    else:
        df = df.sort_values(by=["included", "revenue"], ascending=[False, False], na_position="last")
    return df.reset_index(drop=True)
