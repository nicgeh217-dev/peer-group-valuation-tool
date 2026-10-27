"""Auto-Discovery — screen peer universes from index constituents by GICS sector.

Replaces hand-curated ticker lists with a real screening step, like the first
step in any professional comparable-company analysis: "give me all listed
companies in sector X". Constituents are pulled from a public source (S&P 500
via Wikipedia) and cached locally; the existing size/revenue filter then narrows
to the target's size band.

Adapter-style: fetch_sp500() is one source; a STOXX600 / CapIQ source could be
dropped in the same way.
"""
from __future__ import annotations

import csv
import io

from paths import APP_DIR

CONSTITUENTS_CSV = APP_DIR / "constituents.csv"
_HEADERS = {"User-Agent": "Mozilla/5.0"}

INDEX_SP500 = "S&P 500"
INDEX_STOXX = "STOXX Europe 600"

# European exchange -> Yahoo Finance suffix. Only exchanges that map reliably
# to Yahoo are listed; countries whose Yahoo symbols diverge too often
# (Ireland, Greece, Poland, Luxembourg) are dropped rather than emit dead tickers.
_EU_SUFFIX = {
    "Switzerland": ".SW", "United Kingdom": ".L", "Netherlands": ".AS",
    "France": ".PA", "Germany": ".DE", "Italy": ".MI", "Sweden": ".ST",
    "Spain": ".MC", "Denmark": ".CO", "Finland": ".HE", "Norway": ".OL",
    "Belgium": ".BR", "Austria": ".VI", "Portugal": ".LS",
}
# Well-known STOXX names whose local ticker differs from the Yahoo symbol.
_EU_ALIAS = {
    "Ferrari": "RACE.MI", "Pernod Ricard": "RI.PA", "Swiss Re": "SREN.SW",
    "H&M": "HM-B.ST", "Aker BP": "AKRBP.OL",
}


# ---------------- Sources ----------------

def fetch_sp500() -> list[dict]:
    """Pull current S&P 500 constituents (ticker, name, GICS sector) from Wikipedia."""
    import requests
    import pandas as pd
    r = requests.get(
        "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies",
        headers=_HEADERS, timeout=25,
    )
    r.raise_for_status()
    df = pd.read_html(io.StringIO(r.text))[0]
    out = []
    for _, row in df.iterrows():
        sym = str(row.get("Symbol", "")).strip().replace(".", "-")  # BRK.B → BRK-B (Yahoo)
        if not sym:
            continue
        out.append({
            "ticker": sym,
            "name": str(row.get("Security", "")).strip(),
            "sector": str(row.get("GICS Sector", "")).strip(),
            "region": "US",
            "index": INDEX_SP500,
        })
    return out


def _eu_yahoo_symbol(ticker: str, country: str) -> str | None:
    """Map a STOXX local ticker + country to a Yahoo Finance symbol.
    Handles Nordic share classes ('ATCOa' -> 'ATCO-A', 'AMBU B' -> 'AMBU-B')."""
    suf = _EU_SUFFIX.get(country)
    if not suf:
        return None
    tk = str(ticker).strip()
    if " " in tk:  # 'AMBU B' -> 'AMBU-B'
        base, last = tk.rsplit(" ", 1)
        tk = f"{base.replace(' ', '')}-{last.upper()}" if (len(last) == 1 and last.isalpha()) \
            else tk.replace(" ", "")
    elif tk and tk[-1].islower() and tk[:-1].isupper():  # 'ATCOa' -> 'ATCO-A'
        tk = tk[:-1] + "-" + tk[-1].upper()
    return tk + suf


def fetch_stoxx600() -> list[dict]:
    """Pull STOXX Europe 600 constituents (ticker, name, ICB sector) from Wikipedia,
    mapping each local ticker to a Yahoo Finance symbol via its listing country."""
    import requests
    import pandas as pd
    r = requests.get("https://en.wikipedia.org/wiki/STOXX_Europe_600",
                     headers=_HEADERS, timeout=25)
    r.raise_for_status()
    tables = pd.read_html(io.StringIO(r.text))
    df = next((t for t in tables
               if {"Ticker", "Company", "ICB Sector", "Country"}.issubset(set(map(str, t.columns)))),
              None)
    if df is None:
        return []
    out = []
    for _, row in df.iterrows():
        name = str(row.get("Company", "")).strip()
        sym = _EU_ALIAS.get(name) or _eu_yahoo_symbol(row.get("Ticker", ""), str(row.get("Country", "")).strip())
        if not sym:
            continue
        out.append({
            "ticker": sym,
            "name": name,
            "sector": str(row.get("ICB Sector", "")).strip(),
            "region": "Europe",
            "index": INDEX_STOXX,
        })
    return out


_SOURCES = {INDEX_SP500: fetch_sp500, INDEX_STOXX: fetch_stoxx600}


# ---------------- Cache (constituents.csv) ----------------

def refresh_constituents() -> int:
    """Fetch every source and write constituents.csv. A source that fails is
    skipped so a network hiccup on one index never wipes the others."""
    rows: list[dict] = []
    for fetch in _SOURCES.values():
        try:
            rows += fetch()
        except Exception:
            pass
    if rows:
        with CONSTITUENTS_CSV.open("w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["ticker", "name", "sector", "region", "index"])
            w.writeheader()
            w.writerows(rows)
    return len(rows)


def load_constituents() -> list[dict]:
    """Read constituents.csv; fetch + cache on first use if missing."""
    if not CONSTITUENTS_CSV.exists():
        try:
            refresh_constituents()
        except Exception:
            return []
    out = []
    try:
        with CONSTITUENTS_CSV.open(encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                out.append({k: (v or "").strip() for k, v in row.items()})
    except Exception:
        return []
    return out


def _ensure_index(index: str) -> list[dict]:
    """Constituents for one index; refresh once if that index isn't cached yet
    (e.g. an old cache that predates the STOXX source)."""
    rows = load_constituents()
    if index and not any(c.get("index") == index for c in rows):
        try:
            refresh_constituents()
            rows = load_constituents()
        except Exception:
            pass
    return rows


# ---------------- Screening ----------------

def available_indices() -> list[str]:
    """Indices that actually have cached constituents, in a stable order."""
    rows = _ensure_index(INDEX_STOXX)  # triggers a one-time upgrade fetch if needed
    present = {c.get("index") for c in rows}
    return [ix for ix in (INDEX_SP500, INDEX_STOXX) if ix in present]


def list_gics_sectors(index: str = INDEX_SP500) -> list[str]:
    """Sectors for one index (default S&P 500 / GICS, for backward compatibility
    with the desktop app). STOXX 600 returns ICB sectors."""
    rows = _ensure_index(index)
    return sorted({c["sector"] for c in rows
                   if c.get("sector") and (not index or c.get("index") == index)})


def screen(sector: str | None = None, index: str = INDEX_SP500,
           region: str | None = None, limit: int | None = None) -> list[str]:
    """Return tickers matching the index + sector (+ optional region) filter.
    Defaults to S&P 500 so existing callers (desktop app) are unaffected."""
    out = []
    for c in _ensure_index(index):
        if index and c.get("index") != index:
            continue
        if sector and c.get("sector") != sector:
            continue
        if region and c.get("region") != region:
            continue
        out.append(c["ticker"])
    return out[:limit] if limit else out
