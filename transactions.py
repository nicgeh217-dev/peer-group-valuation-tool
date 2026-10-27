"""Precedent Transactions (M&A deal comps).

ADAPTER PATTERN — same design as data_source.py:
  TransactionSource is the interface. CSVTransactionSource is the prototype
  implementation. In production, drop in a CapIQTransactionSource /
  MergermarketTransactionSource that hits the enterprise database — the
  valuation logic below never changes.

Deal multiples differ from trading multiples: they embed a CONTROL PREMIUM
(acquirers pay 20-40% above the standalone market price for the majority stake).
So precedent-transaction-implied values typically sit ABOVE trading comps.
"""
from __future__ import annotations

import csv
import statistics
from dataclasses import dataclass
from pathlib import Path

from paths import APP_DIR
DEFAULT_CSV = APP_DIR / "transactions.csv"


@dataclass
class Deal:
    date: str | None
    target: str
    acquirer: str | None
    ev: float | None
    revenue: float | None
    ebitda: float | None
    ebit: float | None
    sector: str | None = None

    @property
    def ev_sales(self):
        return self.ev / self.revenue if (self.ev and self.revenue) else None

    @property
    def ev_ebitda(self):
        return self.ev / self.ebitda if (self.ev and self.ebitda) else None

    @property
    def ev_ebit(self):
        return self.ev / self.ebit if (self.ev and self.ebit) else None


# ---------------- Adapter interface ----------------

class TransactionSource:
    """Interface. Implement load() to return a list[Deal]."""
    def load(self) -> list[Deal]:
        raise NotImplementedError


class CSVTransactionSource(TransactionSource):
    """Prototype source: reads deals from a CSV.

    Expected columns (header row, case-insensitive):
      date, target, acquirer, ev, revenue, ebitda, ebit
    Money columns in millions OR absolute — interpreted as-is (keep consistent
    with the target you compare against).
    """
    def __init__(self, path: Path | str = DEFAULT_CSV, scale: float = 1.0):
        self.path = Path(path)
        self.scale = scale  # multiply money columns (e.g. 1e6 if CSV is in millions)

    def load(self) -> list[Deal]:
        if not self.path.exists():
            return []
        deals = []
        with self.path.open(encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            # normalize header keys
            for raw in reader:
                row = {(k or "").strip().lower(): (v or "").strip() for k, v in raw.items()}

                def num(key):
                    v = row.get(key, "")
                    if not v:
                        return None
                    try:
                        return float(v.replace(",", "").replace(" ", "")) * self.scale
                    except ValueError:
                        return None

                target = row.get("target") or row.get("target company")
                if not target:
                    continue
                deals.append(Deal(
                    date=row.get("date") or None,
                    target=target,
                    acquirer=row.get("acquirer") or None,
                    ev=num("ev"),
                    revenue=num("revenue") or num("sales"),
                    ebitda=num("ebitda"),
                    ebit=num("ebit"),
                    sector=row.get("sector") or None,
                ))
        return deals


class FMPTransactionSource(TransactionSource):
    """Live API adapter: Financial Modeling Prep M&A feed.

    Proves the architecture is swappable — same Deal interface, different source.
    NOTE: FMP's free M&A endpoint returns deal ANNOUNCEMENTS (acquirer, target,
    date, SEC link) but NOT deal value or target financials. So deals come back
    with ev/revenue/ebitda = None and won't produce multiples until enriched
    from a paid tier (or CapIQ/Mergermarket). The wiring, however, is real:
    swap CSV → FMP via config and the rest of the pipeline is untouched.
    """
    BASE = "https://financialmodelingprep.com/stable"

    def __init__(self, api_key: str, pages: int = 1):
        self.api_key = api_key
        self.pages = pages

    def load(self) -> list[Deal]:
        import requests
        if not self.api_key:
            return []
        deals = []
        for page in range(self.pages):
            try:
                r = requests.get(f"{self.BASE}/mergers-acquisitions-latest",
                                 params={"apikey": self.api_key, "page": page}, timeout=20)
                if r.status_code != 200:
                    break
                for d in r.json():
                    deals.append(Deal(
                        date=d.get("transactionDate"),
                        target=d.get("targetedCompanyName") or d.get("targetedSymbol") or "?",
                        acquirer=d.get("companyName"),
                        ev=None,        # not provided by free feed
                        revenue=None,
                        ebitda=None,
                        ebit=None,
                        sector=None,    # feed has no sector tag
                    ))
            except Exception:
                break
        return deals


# Where a production adapter would live:
# class CapIQTransactionSource(TransactionSource):
#     def __init__(self, api_key, sector, region, date_from): ...
#     def load(self) -> list[Deal]: ...   # query CapIQ, map to Deal — has deal values


def get_transaction_source(kind: str = "csv", **kwargs) -> TransactionSource:
    """Factory: pick the transaction data source. Add new adapters here."""
    kind = (kind or "csv").lower()
    if kind == "csv":
        return CSVTransactionSource(kwargs.get("path", DEFAULT_CSV), kwargs.get("scale", 1.0))
    if kind == "fmp":
        return FMPTransactionSource(kwargs.get("api_key", ""), kwargs.get("pages", 1))
    raise ValueError(f"Unknown transaction source: {kind!r}")


# ---------------- Statistics + implied valuation ----------------

MULT_KEYS = ["ev_sales", "ev_ebitda", "ev_ebit"]
MULT_LABELS = {"ev_sales": "EV / Sales", "ev_ebitda": "EV / EBITDA", "ev_ebit": "EV / EBIT"}


def deal_multiple_stats(deals: list[Deal]) -> dict:
    """Median / Q1 / Q3 per deal multiple, dropping non-positive values."""
    out = {}
    for key in MULT_KEYS:
        vals = sorted(v for v in (getattr(d, key) for d in deals) if v and v > 0)
        if not vals:
            continue
        out[key] = {
            "n": len(vals),
            "median": statistics.median(vals),
            "q1": _quantile(vals, 0.25),
            "q3": _quantile(vals, 0.75),
            "min": vals[0],
            "max": vals[-1],
        }
    return out


def _quantile(sorted_vals, q):
    if not sorted_vals:
        return None
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    pos = (len(sorted_vals) - 1) * q
    lo = int(pos)
    frac = pos - lo
    if lo + 1 < len(sorted_vals):
        return sorted_vals[lo] * (1 - frac) + sorted_vals[lo + 1] * frac
    return sorted_vals[lo]


def implied_from_transactions(target: dict, stats: dict) -> list[dict]:
    """target: {revenue, ebitda, ebit, net_debt}. Returns implied EV+Equity per multiple."""
    net_debt = target.get("net_debt") or 0
    operand = {"ev_sales": "revenue", "ev_ebitda": "ebitda", "ev_ebit": "ebit"}
    rows = []
    for key in MULT_KEYS:
        s = stats.get(key)
        operand_val = target.get(operand[key])
        if not s or not operand_val:
            continue
        low_ev = operand_val * s["q1"] if s.get("q1") else None
        base_ev = operand_val * s["median"]
        high_ev = operand_val * s["q3"] if s.get("q3") else None
        rows.append({
            "multiple": key,
            "label": MULT_LABELS[key],
            "operand": operand[key],
            "operand_val": operand_val,
            "n": s["n"],
            "q1": s.get("q1"), "median": s["median"], "q3": s.get("q3"),
            "low_ev": low_ev, "base_ev": base_ev, "high_ev": high_ev,
            "low_equity": (low_ev - net_debt) if low_ev else None,
            "base_equity": base_ev - net_debt,
            "high_equity": (high_ev - net_debt) if high_ev else None,
        })
    return rows


def filter_by_sector(deals: list[Deal], universe_name: str | None) -> list[Deal]:
    """Keep deals whose sector tag matches the selected universe.

    Match logic: the deal's sector keyword appears in the universe name
    (case-insensitive), OR vice versa. Deals without a sector tag are kept
    (treated as universal). If no universe given, return all.
    """
    if not universe_name:
        return deals
    uni = universe_name.lower()
    out = []
    for d in deals:
        if not d.sector:
            out.append(d)          # untagged → always included
            continue
        sec = d.sector.lower().strip()
        # token match: any word of the deal sector appears in the universe name
        tokens = [t for t in sec.replace("/", " ").split() if len(t) > 2]
        if any(t in uni for t in tokens) or sec in uni or uni.split()[0] in sec:
            out.append(d)
    return out


def load_transactions(path: Path | str = DEFAULT_CSV, scale: float = 1.0,
                      universe_name: str | None = None,
                      source: str = "csv", api_key: str = "") -> list[Deal]:
    """Load deals from the chosen source (csv | fmp), then filter by sector."""
    src = get_transaction_source(source, path=path, scale=scale, api_key=api_key)
    deals = src.load()
    if universe_name:
        deals = filter_by_sector(deals, universe_name)
    return deals
