"""Data source adapter.

Two backends, selected via config["data_source"]:
  - "yfinance" (default): free, no API key, global coverage via Yahoo Finance
  - "fmp": Financial Modeling Prep (paid plan recommended — free tier covers only ~5 demo tickers)

Architecture note: both adapters return CompanyFinancials so screening / multiples / export
stay completely independent of the source. To add Capital IQ / Bloomberg later, implement
the same `fetch()` and `screen()` interface.
"""
from __future__ import annotations

import json
import os
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests

from paths import CACHE_DIR
CACHE_DIR.mkdir(exist_ok=True)


@dataclass
class CompanyFinancials:
    symbol: str
    name: str | None
    sector: str | None
    industry: str | None
    country: str | None
    currency: str | None             # original reporting currency
    revenue: float | None
    ebitda: float | None
    ebit: float | None
    net_income: float | None
    net_debt: float | None
    enterprise_value: float | None
    market_cap: float | None
    ebitda_margin: float | None      # ratios — currency-agnostic
    ebit_margin: float | None
    ev_sales: float | None           # multiples — currency-agnostic
    ev_ebitda: float | None
    ev_ebit: float | None
    pe: float | None
    fiscal_date: str | None
    target_currency: str | None = None      # currency the amount fields are in
    fx_rate: float | None = None            # rate used: original -> target
    business_summary: str | None = None     # Yahoo longBusinessSummary
    match_score: int | None = None          # AI business-model match 1-10
    match_rationale: str | None = None      # AI one-line reasoning
    business_profile: str | None = None     # AI structured profile (multi-line)
    beta: float | None = None               # Yahoo levered beta (5y vs market)


# ---------------- yfinance backend ----------------

class YFinanceDataSource:
    """Yahoo Finance via the `yfinance` package.

    Free, no key, global coverage. For non-US tickers use Yahoo suffixes,
    e.g. SAP.DE (Xetra), CAP.PA (Paris), SAP.L (London).
    """

    def __init__(self, throttle_seconds: float = 0.1, target_currency: str = "USD"):
        self.target_currency = (target_currency or "USD").upper()
        try:
            import yfinance as yf  # noqa: F401
        except ModuleNotFoundError as e:
            import sys, site, subprocess
            diag = (
                f"yfinance Import fehlgeschlagen.\n"
                f"sys.executable: {sys.executable}\n"
                f"User site enabled: {site.ENABLE_USER_SITE}\n"
                f"User site dir: {site.getusersitepackages()}\n"
                f"sys.path:\n  " + "\n  ".join(sys.path) + "\n\n"
                f"Versuche jetzt automatische Installation..."
            )
            print(diag)
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "yfinance"])
            # Refresh sys.path to pick up newly installed package
            site.main()
            import yfinance as yf  # noqa: F401
        self.throttle = throttle_seconds
        self._last_call = 0.0
        # fetch() may be called from several threads (the web frontend fetches the
        # universe in parallel). Without the lock the throttle is a read-modify-write
        # race and every thread would hit Yahoo at once.
        self._throttle_lock = threading.Lock()

    def _throttle(self):
        with self._throttle_lock:
            elapsed = time.time() - self._last_call
            if elapsed < self.throttle:
                time.sleep(self.throttle - elapsed)
            self._last_call = time.time()

    def screen(self, **kwargs):
        """yfinance has no screener — caller must pass an explicit ticker list."""
        return None

    def fetch(self, symbol: str) -> CompanyFinancials | None:
        import yfinance as yf
        cache_file = CACHE_DIR / f"yf_{symbol.replace('/', '_')}.json"
        info = None
        if cache_file.exists():
            try:
                info = json.loads(cache_file.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                info = None
        if info is None:
            self._throttle()
            try:
                t = yf.Ticker(symbol)
                info = t.info or {}
            except Exception as e:
                print(f"      yfinance error for {symbol}: {e}")
                return None
            # Only cache non-trivial responses
            if info and info.get("symbol"):
                try:
                    cache_file.write_text(json.dumps(info, default=str), encoding="utf-8")
                except Exception:
                    pass

        if not info or not info.get("symbol"):
            return None

        # Filter out non-equities (indices, ETFs, mutual funds, currencies)
        # — these slip through when a ticker symbol is wrong or ambiguous.
        quote_type = (info.get("quoteType") or "").upper()
        if quote_type and quote_type != "EQUITY":
            print(f"      {symbol}: skipped (quoteType={quote_type})")
            return None

        revenue = info.get("totalRevenue")
        # No revenue = no peer analysis possible. Skip cleanly instead of
        # carrying empty rows into the Long List.
        if revenue is None or revenue == 0:
            print(f"      {symbol}: skipped (no revenue data)")
            return None

        ebitda = info.get("ebitda")
        ev = info.get("enterpriseValue")
        mcap = info.get("marketCap")
        ebitda_margin = info.get("ebitdaMargins")
        op_margin = info.get("operatingMargins")
        net_margin = info.get("profitMargins")

        ebit = revenue * op_margin if (revenue and op_margin is not None) else None
        net_income = revenue * net_margin if (revenue and net_margin is not None) else None

        # Yahoo doesn't always give netDebt; derive from totalDebt - cash if missing
        net_debt = info.get("netDebt")
        if net_debt is None:
            total_debt = info.get("totalDebt")
            cash = info.get("totalCash")
            if total_debt is not None and cash is not None:
                net_debt = total_debt - cash

        ev_sales = info.get("enterpriseToRevenue")
        ev_ebitda = info.get("enterpriseToEbitda")
        ev_ebit = (ev / ebit) if (ev and ebit) else None
        pe = info.get("trailingPE")

        # ---- Currency normalization ----
        # Multiples and margins are dimensionless and need no conversion.
        # Only revenue / EBITDA / EBIT / NI / EV / Net Debt / Market Cap.
        from fx import get_rate, convert as fx_convert
        original_ccy = info.get("financialCurrency") or info.get("currency")
        rate = get_rate(original_ccy, self.target_currency) if original_ccy else None
        if rate is not None:
            revenue = fx_convert(revenue, original_ccy, self.target_currency)
            ebitda = fx_convert(ebitda, original_ccy, self.target_currency)
            ebit = fx_convert(ebit, original_ccy, self.target_currency)
            net_income = fx_convert(net_income, original_ccy, self.target_currency)
            net_debt = fx_convert(net_debt, original_ccy, self.target_currency)
            ev = fx_convert(ev, original_ccy, self.target_currency)
            mcap = fx_convert(mcap, original_ccy, self.target_currency)

        return CompanyFinancials(
            symbol=symbol,
            name=info.get("longName") or info.get("shortName"),
            sector=info.get("sector"),
            industry=info.get("industry"),
            country=info.get("country"),
            currency=original_ccy,
            target_currency=self.target_currency,
            fx_rate=rate,
            business_summary=info.get("longBusinessSummary"),
            beta=info.get("beta"),
            revenue=revenue,
            ebitda=ebitda,
            ebit=ebit,
            net_income=net_income,
            net_debt=net_debt,
            enterprise_value=ev,
            market_cap=mcap,
            ebitda_margin=ebitda_margin,
            ebit_margin=op_margin,
            ev_sales=ev_sales,
            ev_ebitda=ev_ebitda,
            ev_ebit=ev_ebit,
            pe=pe,
            fiscal_date=info.get("mostRecentQuarter") or None,
        )


# ---------------- FMP backend ----------------

class FMPDataSource:
    BASE = "https://financialmodelingprep.com/stable"

    def __init__(self, api_key: str, throttle_seconds: float = 0.25):
        if not api_key or api_key.startswith("YOUR_"):
            raise ValueError("Missing FMP API key.")
        self.api_key = api_key
        self.throttle = throttle_seconds
        self._last_call = 0.0

    def _get(self, url: str, params: dict[str, Any]) -> Any:
        params = {k: v for k, v in params.items() if v is not None}
        params["apikey"] = self.api_key
        cache_key_params = {k: v for k, v in params.items() if k != "apikey"}
        cache_key = url.replace("https://", "").replace("/", "_") + "_" + json.dumps(
            cache_key_params, sort_keys=True
        )
        cache_key = "".join(c if c.isalnum() or c in "._-" else "_" for c in cache_key)[:200]
        cache_file = CACHE_DIR / f"{cache_key}.json"
        if cache_file.exists():
            try:
                return json.loads(cache_file.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                pass
        elapsed = time.time() - self._last_call
        if elapsed < self.throttle:
            time.sleep(self.throttle - elapsed)
        self._last_call = time.time()
        resp = requests.get(url, params=params, timeout=30)
        if resp.status_code == 429:
            time.sleep(2.0)
            resp = requests.get(url, params=params, timeout=30)
        if resp.status_code in (402, 403):
            return {"_paywalled": True, "status": resp.status_code}
        resp.raise_for_status()
        data = resp.json()
        cache_file.write_text(json.dumps(data), encoding="utf-8")
        return data

    def screen(self, sector=None, industry=None, countries=None,
               market_cap_min=None, market_cap_max=None, limit=100):
        url = f"{self.BASE}/company-screener"
        seen: dict[str, dict] = {}
        any_ok = False
        for country in (countries or [None]):
            params = {
                "sector": sector, "industry": industry, "country": country,
                "marketCapMoreThan": market_cap_min, "marketCapLowerThan": market_cap_max,
                "isActivelyTrading": "true", "limit": limit,
            }
            data = self._get(url, params)
            if isinstance(data, dict) and data.get("_paywalled"):
                return None
            any_ok = True
            if not isinstance(data, list):
                continue
            for row in data:
                sym = row.get("symbol")
                if sym and sym not in seen:
                    seen[sym] = row
        return list(seen.values()) if any_ok else None

    def fetch(self, symbol: str) -> CompanyFinancials | None:
        prof = self._get(f"{self.BASE}/profile", {"symbol": symbol})
        km = self._get(f"{self.BASE}/key-metrics-ttm", {"symbol": symbol})
        ratios = self._get(f"{self.BASE}/ratios-ttm", {"symbol": symbol})
        bs = self._get(f"{self.BASE}/balance-sheet-statement", {"symbol": symbol, "limit": 1})

        def first(x):
            if isinstance(x, dict) and x.get("_paywalled"):
                return {}
            if isinstance(x, list) and x:
                return x[0]
            return {}

        prof_r, km_r, rat_r, bs_r = first(prof), first(km), first(ratios), first(bs)
        if not prof_r and not km_r:
            return None

        mcap = km_r.get("marketCap") or prof_r.get("mktCap") or prof_r.get("marketCap")
        ev = km_r.get("enterpriseValueTTM") or km_r.get("enterpriseValue")
        rev_ps = km_r.get("revenuePerShareTTM")
        shares = km_r.get("sharesOutstandingTTM") or prof_r.get("sharesOutstanding")
        revenue = rev_ps * shares if (rev_ps and shares) else None

        ebitda_margin = rat_r.get("ebitdaMarginTTM")
        ebit_margin = rat_r.get("ebitMarginTTM")
        net_margin = rat_r.get("netProfitMarginTTM")
        ebitda = revenue * ebitda_margin if (revenue and ebitda_margin is not None) else None
        ebit = revenue * ebit_margin if (revenue and ebit_margin is not None) else None
        net_income = revenue * net_margin if (revenue and net_margin is not None) else None

        total_debt = bs_r.get("totalDebt")
        cash = bs_r.get("cashAndShortTermInvestments") or bs_r.get("cashAndCashEquivalents")
        net_debt = total_debt - cash if (total_debt is not None and cash is not None) else None

        ev_sales = km_r.get("evToSalesTTM")
        ev_ebitda = km_r.get("enterpriseValueOverEBITDATTM") or km_r.get("evToEBITDATTM")
        ev_ebit = (ev / ebit) if (ev and ebit) else None
        pe = rat_r.get("priceEarningsRatioTTM") or rat_r.get("peRatioTTM")

        return CompanyFinancials(
            symbol=symbol, name=prof_r.get("companyName"),
            sector=prof_r.get("sector"), industry=prof_r.get("industry"),
            country=prof_r.get("country"), currency=prof_r.get("currency"),
            revenue=revenue, ebitda=ebitda, ebit=ebit, net_income=net_income,
            net_debt=net_debt, enterprise_value=ev, market_cap=mcap,
            ebitda_margin=ebitda_margin, ebit_margin=ebit_margin,
            ev_sales=ev_sales, ev_ebitda=ev_ebitda, ev_ebit=ev_ebit, pe=pe,
            fiscal_date=bs_r.get("date"),
        )


def load_data_source(config: dict):
    backend = (config.get("data_source") or "yfinance").lower()
    if backend == "yfinance":
        return YFinanceDataSource()
    if backend == "fmp":
        key = config.get("fmp_api_key") or os.environ.get("FMP_API_KEY", "")
        return FMPDataSource(api_key=key)
    raise ValueError(f"Unknown data_source: {backend!r}")
