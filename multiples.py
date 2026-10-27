"""Outlier detection and summary stats. Multiples now come pre-computed from the data source."""
from __future__ import annotations

import pandas as pd

MULTIPLE_COLS = ["ev_sales", "ev_ebitda", "ev_ebit", "pe"]


def flag_outliers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    flags = [[] for _ in range(len(df))]

    for col in MULTIPLE_COLS:
        if col not in df.columns:
            continue
        s = pd.to_numeric(df[col], errors="coerce")
        for i, v in enumerate(s):
            if v is not None and not pd.isna(v) and v < 0:
                flags[i].append(f"{col}<0")
        pos = s[s > 0].dropna()
        if len(pos) >= 4:
            q1, q3 = pos.quantile(0.25), pos.quantile(0.75)
            iqr = q3 - q1
            lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            for i, v in enumerate(s):
                if v is None or pd.isna(v) or v <= 0:
                    continue
                if v < lo or v > hi:
                    flags[i].append(f"{col}_outlier")

    df["outlier_flags"] = ["; ".join(f) if f else "" for f in flags]
    df["is_outlier"] = df["outlier_flags"] != ""
    return df


def summary_stats(df: pd.DataFrame, exclude_outliers: bool = True) -> pd.DataFrame:
    if exclude_outliers and "is_outlier" in df.columns:
        clean = df[~df["is_outlier"]]
    else:
        clean = df

    rows = []
    for col in MULTIPLE_COLS + ["ebitda_margin", "ebit_margin"]:
        if col not in clean.columns:
            continue
        s = pd.to_numeric(clean[col], errors="coerce").dropna()
        if col in MULTIPLE_COLS:
            s = s[s > 0]
        if s.empty:
            rows.append({"metric": col, "n": 0, "mean": None, "median": None,
                         "q1": None, "q3": None, "min": None, "max": None})
            continue
        rows.append({
            "metric": col,
            "n": int(s.count()),
            "mean": float(s.mean()),
            "median": float(s.median()),
            "q1": float(s.quantile(0.25)),
            "q3": float(s.quantile(0.75)),
            "min": float(s.min()),
            "max": float(s.max()),
        })
    return pd.DataFrame(rows)
