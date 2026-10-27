"""FX conversion using Yahoo Finance currency pairs.

Pair convention: '{FROM}{TO}=X' (e.g. EURUSD=X = EUR→USD multiplier).
Same-currency pairs return 1.0 without a network call.
"""
from __future__ import annotations

import json
import threading
import time
from pathlib import Path

from paths import CACHE_DIR
CACHE_FILE = CACHE_DIR / "fx_rates.json"
CACHE_FILE.parent.mkdir(exist_ok=True)

# In-process cache so we don't re-read the JSON on every fetch
_RATES: dict[str, float] = {}
_LOADED = False
# get_rate() is called from the parallel fetch threads. _RATES and the cache file
# are process-global, so both the lookup-miss path and the file write need a lock,
# otherwise two threads fetch the same pair and can write a truncated JSON file.
_LOCK = threading.RLock()

# Common currency aliases — Yahoo uses ISO codes but some sources differ
ALIASES = {
    "GBp": "GBP",   # Yahoo uses GBp (pence) for some UK reports
    "ILA": "ILS",
    "ZAc": "ZAR",
}


def _normalize(ccy: str | None) -> str | None:
    if not ccy:
        return None
    ccy = ccy.strip().upper()
    return ALIASES.get(ccy, ccy)


def _load_cache():
    global _LOADED, _RATES
    if _LOADED:
        return
    # guarded by _LOCK from get_rate; direct callers are single-threaded
    if CACHE_FILE.exists():
        try:
            data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            # Refuse rates older than 24h
            if time.time() - data.get("_timestamp", 0) < 86400:
                _RATES = data.get("rates", {})
        except Exception:
            _RATES = {}
    _LOADED = True


def _save_cache():
    try:
        CACHE_FILE.write_text(
            json.dumps({"_timestamp": time.time(), "rates": _RATES}, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def get_rate(from_ccy: str, to_ccy: str) -> float | None:
    """Returns multiplier to convert FROM → TO. None if not resolvable."""
    f, t = _normalize(from_ccy), _normalize(to_ccy)
    if not f or not t:
        return None
    if f == t:
        return 1.0

    key = f"{f}{t}"
    with _LOCK:
        _load_cache()
        if key in _RATES:
            return _RATES[key]
        return _fetch_and_cache(f, t, key)


def _fetch_and_cache(f: str, t: str, key: str) -> float | None:
    """Network path of get_rate. Caller holds _LOCK."""

    # Try direct pair first, then inverse, then via USD
    try:
        import yfinance as yf

        rate = _fetch_pair(yf, f, t)
        if rate is None and f != "USD" and t != "USD":
            # Triangulate via USD: f→USD * USD→t
            f_usd = _fetch_pair(yf, f, "USD")
            usd_t = _fetch_pair(yf, "USD", t)
            if f_usd is not None and usd_t is not None:
                rate = f_usd * usd_t

        if rate is not None:
            _RATES[key] = rate
            _save_cache()
        return rate
    except Exception:
        return None


def _fetch_pair(yf_module, f: str, t: str) -> float | None:
    """Try {f}{t}=X (direct) and if missing the inverse {t}{f}=X."""
    for sym, invert in [(f"{f}{t}=X", False), (f"{t}{f}=X", True)]:
        try:
            info = yf_module.Ticker(sym).info or {}
            price = info.get("regularMarketPrice") or info.get("previousClose")
            if price and price > 0:
                return 1.0 / price if invert else float(price)
        except Exception:
            continue
    return None


def convert(value, from_ccy: str | None, to_ccy: str) -> float | None:
    """Convert a numeric value. Returns None if value is None or no rate found."""
    if value is None:
        return None
    rate = get_rate(from_ccy, to_ccy) if from_ccy else None
    if rate is None:
        return None
    try:
        return float(value) * rate
    except (TypeError, ValueError):
        return None
