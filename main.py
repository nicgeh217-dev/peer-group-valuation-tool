"""Orchestration. Supports two input modes:

1. **tickers**: explicit candidate list (works on FMP Free Tier)
2. **screen**: auto-screening via FMP company-screener (requires paid tier)
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from data_source import load_data_source
from export import export
from multiples import flag_outliers, summary_stats
from screening import apply_filters, rank_by_similarity


def run(config_path: Path, output_path: Path | None = None) -> Path:
    config = json.loads(Path(config_path).read_text(encoding="utf-8"))
    target = config["target"]
    limits = config.get("limits", {})
    max_candidates = limits.get("max_candidates", 80)

    ds = load_data_source(config)

    # ---- 1) Build candidate list ----
    tickers: list[str] = []
    explicit = config.get("tickers") or target.get("tickers")
    if explicit:
        tickers = list(explicit)
        print(f"[1/4] Using explicit ticker list: {len(tickers)} candidates.")
    else:
        print(f"[1/4] Trying screener (sector={target.get('sector')!r})...")
        countries = target.get("country")
        if isinstance(countries, str):
            countries = [countries]
        hits = ds.screen(
            sector=target.get("sector"),
            industry=target.get("industry"),
            countries=countries,
            market_cap_min=target.get("market_cap_min_usd"),
            market_cap_max=target.get("market_cap_max_usd"),
            limit=max_candidates,
        )
        if hits is None:
            print("      Screener not available on your FMP plan (Free tier).")
            print("      Provide a candidate ticker list via 'tickers' in config.json.")
            sys.exit(2)
        tickers = [h.get("symbol") for h in hits if h.get("symbol")]
        print(f"      Screener returned {len(tickers)} tickers.")

    tickers = tickers[:max_candidates]

    # ---- 2) Fetch financials ----
    print(f"[2/4] Fetching financials for {len(tickers)} tickers...")
    rows = []
    for i, sym in enumerate(tickers, 1):
        try:
            fin = ds.fetch(sym)
        except Exception as e:
            print(f"      [{i}/{len(tickers)}] {sym}: error — {e}")
            continue
        if fin is None:
            print(f"      [{i}/{len(tickers)}] {sym}: no data")
            continue
        rows.append(fin.__dict__)
        if i % 5 == 0 or i == len(tickers):
            print(f"      [{i}/{len(tickers)}] processed")

    if not rows:
        print("No financials retrieved.")
        sys.exit(1)
    df = pd.DataFrame(rows)

    # ---- 3) Filters + outliers + ranking ----
    print(f"[3/4] Applying filters and detecting outliers...")
    df = apply_filters(
        df,
        revenue_min=target.get("revenue_min_usd"),
        revenue_max=target.get("revenue_max_usd"),
        ebitda_margin_min=target.get("ebitda_margin_min"),
        ebitda_margin_max=target.get("ebitda_margin_max"),
    )
    df = flag_outliers(df)

    rev_mid = None
    rmin, rmax = target.get("revenue_min_usd"), target.get("revenue_max_usd")
    if rmin and rmax:
        rev_mid = (rmin + rmax) / 2
    df = rank_by_similarity(df, target_revenue=rev_mid)

    summary = summary_stats(df[df["included"]], exclude_outliers=True)

    # ---- 4) Excel ----
    print(f"[4/4] Writing Excel...")
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_name = (target.get("name") or "peer_group").replace(" ", "_")
        output_path = Path(__file__).parent / "output" / f"{target_name}_{ts}.xlsx"

    countries_meta = target.get("country")
    if isinstance(countries_meta, list):
        countries_meta = ", ".join(countries_meta)
    target_meta = {
        "Target": target.get("name"),
        "Sector": target.get("sector"),
        "Industry": target.get("industry"),
        "Countries": countries_meta or "",
        "Revenue range (USD)": f"{target.get('revenue_min_usd')} – {target.get('revenue_max_usd')}",
        "Market cap range (USD)": f"{target.get('market_cap_min_usd')} – {target.get('market_cap_max_usd')}",
        "Candidates evaluated": len(df),
        "Included after filters": int(df["included"].sum()),
        "Outliers flagged": int(df["is_outlier"].sum()),
        "Mode": "explicit tickers" if explicit else "screener",
        "Run timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    path = export(df, summary, target_meta, output_path)
    print(f"\nDone. Excel: {path}")
    print(f"  Included: {target_meta['Included after filters']} | "
          f"Outliers: {target_meta['Outliers flagged']}")
    return path


def main():
    ap = argparse.ArgumentParser(description="Peer Group Screening Tool")
    ap.add_argument("--config", "-c", required=True, help="Path to config.json")
    ap.add_argument("--output", "-o", help="Output Excel path (optional)")
    args = ap.parse_args()
    run(Path(args.config), Path(args.output) if args.output else None)


if __name__ == "__main__":
    main()
