"""Peer Group Screening — Streamlit web frontend.

Wraps the SAME backend as the desktop GUI (gui.py) with full feature parity:
Trading Comps · DCF / WACC · Precedent Transactions (CSV or FMP live) ·
AI Peer Match · editable ticker universe · Short List · settings persistence ·
per-run output folder · Excel + PowerPoint export.

Run with:  streamlit run app.py
"""
from __future__ import annotations

import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data_source import YFinanceDataSource
from screening import apply_filters, rank_by_similarity
from multiples import flag_outliers, summary_stats
from valuation import compute_implied
from wacc import compute_wacc, dcf, dcf_sensitivity, fcf_bridge
from transactions import (Deal, load_transactions, deal_multiple_stats,
                          implied_from_transactions)
from export import export
from universes import list_universes, get_universe
import discovery
import ai
from paths import APP_DIR, OUTPUT_DIR, CACHE_DIR

# ---------------------------------------------------------------- constants
APP_VERSION = "2.1"
AUTO_PREFIX = "🔍 "  # canonical auto universe key (never rendered; keeps saved
                    # settings_web.json / settings.json values parseable)
# Chart palette — same tokens as assets/style.css and .streamlit/config.toml.
INK = "#14181D"
ACCENT = "#0F5C4B"
SLATE = "#6E7A86"
RULE = "#E2E4E8"
MUTED = "#767C86"
SCALE = 1e6
WEB_SETTINGS = APP_DIR / "settings_web.json"
DEMO_CASE = APP_DIR / "demo_case.json"          # curated inputs for a live demo
SNAPSHOT = CACHE_DIR / "last_run_web.json"      # last completed run, for offline demo
BASELINE = CACHE_DIR / "baseline_run_web.json"  # pinned run to compare against
# Parallel Yahoo requests. 6 keeps the run well under 5s for a typical universe
# without tripping rate limits; the per-source throttle still applies on top.
FETCH_WORKERS = 6

TX_CSV_LABEL = "CSV file"
TX_FMP_LABEL = "FMP live API"

SOURCE_CURATED = "Curated peer sets"
# auto-screen source label per index (only shown if the index has data)
AUTO_LABELS = {discovery.INDEX_SP500: "S&P 500 (US)",
               discovery.INDEX_STOXX: "STOXX 600 (Europe)"}
LABEL_TO_INDEX = {v: k for k, v in AUTO_LABELS.items()}
AUTO_CAPTION = {discovery.INDEX_SP500: "Live-screened from S&P 500 constituents · US large-cap.",
                discovery.INDEX_STOXX: "Live-screened from STOXX Europe 600 · European large & mid cap."}

# Same keys as the desktop app's settings.json (gui.py _collect_settings)
SETTING_KEYS = [
    "target_name", "universe", "currency",
    "rev_min", "rev_max", "mar_min", "mar_max",
    "t_revenue", "t_ebitda", "t_ebit", "t_net_income", "t_net_debt",
    "w_rf", "w_mrp", "w_tax", "w_kd", "w_target_de", "w_size_prem",
    "d_base_fcf", "d_growth", "d_term_g", "d_capex_pct", "d_nwc_pct",
    "tx_csv", "tx_source", "make_pptx", "target_desc", "tickers", "use_ai",
]

DEFAULTS = {
    "target_name": "My_Peer_Group",
    "currency": "USD",
    "rev_min": "200", "rev_max": "5000",
    "mar_min": "", "mar_max": "",
    "t_revenue": "", "t_ebitda": "", "t_ebit": "", "t_net_income": "", "t_net_debt": "",
    "w_rf": "2.5", "w_mrp": "5.5", "w_tax": "25", "w_kd": "5.0",
    "w_target_de": "0.5", "w_size_prem": "0",
    "d_base_fcf": "", "d_growth": "8", "d_term_g": "2.5",
    "d_capex_pct": "5", "d_nwc_pct": "10",
    "tx_csv": "transactions.csv",
    "tx_source": TX_CSV_LABEL,
    "make_pptx": True,
    "target_desc": "",
    "use_ai": True,
}

st.set_page_config(page_title="Peer Group Screening", page_icon="◆", layout="wide")


# ---------------------------------------------------------------- helpers
@st.cache_data(ttl=3600, show_spinner=False)
def avail_indices() -> list[str]:
    try:
        return discovery.available_indices()
    except Exception:
        return [discovery.INDEX_SP500]


@st.cache_data(ttl=3600, show_spinner=False)
def sectors_for(index: str) -> list[str]:
    try:
        return discovery.list_gics_sectors(index)
    except Exception:
        return []


def auto_canonical(index: str, sector: str) -> str:
    return f"{AUTO_PREFIX}{index} · {sector}"


def parse_auto(u: str) -> tuple[str, str] | None:
    """'🔍 STOXX Europe 600 · Banks' -> ('STOXX Europe 600', 'Banks')."""
    if not u.startswith(AUTO_PREFIX):
        return None
    body = u[len(AUTO_PREFIX):]
    return tuple(body.split(" · ", 1)) if " · " in body else None


@st.cache_data(ttl=3600, show_spinner=False)
def universe_options() -> list[str]:
    """All canonical universe strings (curated keys + auto index·sector), for
    settings validation and the first-load seed."""
    opts = list(list_universes())
    for ix in avail_indices():
        opts += [auto_canonical(ix, s) for s in sectors_for(ix)]
    return opts


def tickers_for(u: str) -> list[str]:
    p = parse_auto(u)
    if p:
        return discovery.screen(sector=p[1], index=p[0])
    return get_universe(u)


def clean_target_name(u: str) -> str:
    p = parse_auto(u)
    if p:
        idx = p[0].replace("S&P 500", "SP500").replace("STOXX Europe 600", "STOXX600")
        return f"{idx}_{p[1]}".replace(" ", "_").replace("&", "")
    return u.split("—")[0].strip().replace(" / ", "_").replace(" ", "_")


def pretty_universe(key: str) -> str:
    """Display label for a curated universe key: 'Name — Region/…' -> 'Name (Region)'.
    Display only; the stored value stays the raw dict key."""
    if " — " not in key:
        return key
    name, region = key.split(" — ", 1)
    region = region.split("/")[0].split("(")[0].strip()  # 'Europe/DACH'->'Europe', 'US (broad)'->'US'
    return f"{name} ({region})"


def universe_to_widgets(canonical: str) -> tuple[str, str]:
    """Canonical universe string -> (source label, sub-selection)."""
    p = parse_auto(canonical)
    if p and p[0] in AUTO_LABELS:
        return AUTO_LABELS[p[0]], p[1]
    return SOURCE_CURATED, canonical


def parse_tickers(raw: str) -> list[str]:
    parts = [p.strip() for chunk in raw.splitlines() for p in chunk.split(",")]
    seen, out = set(), []
    for p in parts:
        if p and p.upper() not in seen:
            seen.add(p.upper())
            out.append(p)
    return out


def mio(x):
    try:
        return float(str(x).replace(",", ".")) * SCALE if str(x).strip() else None
    except ValueError:
        return None


def pct(x):
    try:
        return float(str(x).replace(",", ".").rstrip("%")) / 100 if str(x).strip() else None
    except ValueError:
        return None


def style_fig(fig, height: int, title: str = "", **layout):
    """One chart language for the whole app: hairline grid, mono tick labels,
    no chart junk. Keeps every figure readable when projected in a meeting."""
    base = dict(
        height=height,
        template="none",
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="IBM Plex Mono, Consolas, monospace", size=11, color=MUTED),
        title=dict(text=title,
                   font=dict(family="IBM Plex Sans, Segoe UI, sans-serif",
                             size=13, color=INK), x=0, xanchor="left", pad=dict(b=8)),
        margin=dict(l=8, r=14, t=42 if title else 12, b=8),
        showlegend=False,
        hoverlabel=dict(bgcolor=INK, bordercolor=INK, font=dict(
            family="IBM Plex Mono, Consolas, monospace", size=11, color="#FFFFFF")),
    )
    base.update(layout)  # caller wins, so per-chart margins/axes stay overridable
    fig.update_layout(**base)
    fig.update_xaxes(showgrid=True, gridcolor=RULE, gridwidth=1, zeroline=False,
                     linecolor=RULE, ticks="outside", tickcolor=RULE, ticklen=4,
                     title_font=dict(size=11, color=MUTED))
    # automargin: category labels on horizontal bar charts would otherwise be
    # painted outside the plot area and clipped by the tight left margin.
    fig.update_yaxes(showgrid=False, zeroline=False, linecolor=RULE, automargin=True,
                     ticks="outside", tickcolor=RULE, ticklen=4,
                     tickfont=dict(family="IBM Plex Sans, Segoe UI, sans-serif",
                                   size=12, color=INK),
                     title_font=dict(size=11, color=MUTED))
    return fig


def target_sanity(t: dict) -> list[str]:
    """Plausibility checks on the target financials. Ratios cancel the currency
    scale, so values may be raw or in millions. Returns human-readable warnings."""
    rev, eb, ebit, ni = (t.get("revenue"), t.get("ebitda"),
                         t.get("ebit"), t.get("net_income"))
    warn = []
    for label, v in (("Revenue", rev), ("EBITDA", eb), ("EBIT", ebit)):
        if v is not None and v < 0:
            warn.append(f"{label} is negative. Check the sign.")
    if rev and rev > 0 and eb is not None and eb > rev:
        warn.append(f"EBITDA margin > 100% ({eb/rev:.0%}) is impossible. EBITDA cannot exceed revenue, "
                    "Likely a typo in the EBITDA field.")
    elif rev and rev > 0 and eb and (eb / rev) > 0.6:
        warn.append(f"EBITDA margin {eb/rev:.0%} is unusually high. Double-check the figure.")
    if eb and ebit and ebit > eb * 1.001:
        warn.append(f"EBIT ({ebit/1e6:,.0f}) > EBITDA ({eb/1e6:,.0f}). EBIT is after "
                    "D&A, so it should be lower. Check both fields.")
    if eb and ni and ni > eb * 1.001:
        warn.append(f"Net income above EBITDA is possible only with large one-offs; verify.")
    return warn


def fmp_key() -> str:
    """FMP API key from env or config.json (same lookup as the desktop app)."""
    k = os.environ.get("FMP_API_KEY", "")
    if k:
        return k
    cfg = APP_DIR / "config.json"
    if cfg.exists():
        try:
            return json.loads(cfg.read_text(encoding="utf-8")).get("fmp_api_key", "") or ""
        except Exception:
            pass
    return ""


# ---------------------------------------------------------------- settings persistence
def load_web_settings():
    """Seed session_state from settings_web.json — once per browser session,
    BEFORE any widget is instantiated."""
    if st.session_state.get("_settings_loaded"):
        return
    st.session_state["_settings_loaded"] = True
    for k, v in DEFAULTS.items():
        st.session_state.setdefault(k, v)
    if not WEB_SETTINGS.exists():
        return
    try:
        s = json.loads(WEB_SETTINGS.read_text(encoding="utf-8"))
    except Exception:
        return
    if s.get("universe") not in universe_options():
        s.pop("universe", None)
    # normalise tx_source (accept desktop-style "fmp  (…)" strings too)
    if str(s.get("tx_source", "")).lower().startswith("fmp"):
        s["tx_source"] = TX_FMP_LABEL
    elif s.get("tx_source"):
        s["tx_source"] = TX_CSV_LABEL
    for k in SETTING_KEYS:
        if k in s and s[k] is not None:
            st.session_state[k] = s[k]


def save_web_settings():
    try:
        data = {k: st.session_state.get(k) for k in SETTING_KEYS}
        WEB_SETTINGS.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------- demo case
def load_demo_case() -> bool:
    """Overwrite the inputs with the saved demo case (demo_case.json).

    Kept in its own file because settings_web.json is rewritten on every run, so a
    demo that has to work in a meeting must not live there.

    Must run BEFORE the sidebar widgets are created: Streamlit refuses writes to
    session_state for a key whose widget already exists in this script run. The
    button therefore only raises a flag and reruns.
    """
    if not DEMO_CASE.exists():
        return False
    try:
        s = json.loads(DEMO_CASE.read_text(encoding="utf-8"))
    except Exception:
        return False
    for k in SETTING_KEYS:
        if k in s and s[k] is not None:
            st.session_state[k] = s[k]
    # re-seed the two-step universe picker from the loaded canonical value
    src, sub = universe_to_widgets(st.session_state.get("universe", ""))
    st.session_state["uni_source"] = src
    if src == SOURCE_CURATED:
        st.session_state["uni_curated"] = sub
    else:
        st.session_state["uni_sector"] = sub
    return True


# ---------------------------------------------------------------- run snapshot
def _json_safe(o):
    """numpy / pandas scalars are not JSON-serialisable out of the box."""
    if hasattr(o, "item"):          # numpy scalar
        return o.item()
    if isinstance(o, (pd.Timestamp, datetime)):
        return o.isoformat()
    if pd.isna(o):
        return None
    return str(o)


def _run_payload(res: dict) -> dict:
    """Serialisable form of a result bundle. Shared by the snapshot and the
    pinned baseline so both files stay readable by the same loader."""
    prec = res.get("precedent")
    return {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "target_name": st.session_state.get("target_name"),
        "universe": st.session_state.get("universe"),
        "df": res["df"].to_dict("records"),
        "summary": res["summary"].to_dict("records"),
        "implied": res.get("implied"),
        "target": res.get("target"),
        "has_target": res.get("has_target"),
        "wacc": res.get("wacc"),
        "dcf": res.get("dcf"),
        "dcf_sens": res.get("dcf_sens"),
        "convergence": res.get("convergence"),
        "precedent": ({"deals": [asdict(d) for d in prec["deals"]],
                       "included": prec.get("included"),
                       "stats": prec["stats"], "implied": prec["implied"]}
                      if prec else None),
        "log": res.get("log"),
        "sanity": res.get("sanity"),
        "ccy": res.get("ccy"),
        "n_tickers": res.get("n_tickers"),
        "skipped": res.get("skipped"),
        "is_shortlist": res.get("is_shortlist"),
        "xlsx_path": res.get("xlsx_path"),
        "pptx_path": res.get("pptx_path"),
    }


def snapshot_save(res: dict) -> None:
    """Persist the last completed run to cache/last_run_web.json.

    Purpose is demo safety: the live data sources are free tiers without an SLA, so
    a meeting must not depend on them. Plain JSON rather than pickle so the file
    stays inspectable and loading it cannot execute code.
    """
    try:
        SNAPSHOT.write_text(json.dumps(_run_payload(res), ensure_ascii=False, indent=2,
                                       default=_json_safe), encoding="utf-8")
    except Exception:
        pass  # a failed snapshot must never break a successful run


def baseline_save(res: dict) -> bool:
    """Pin the current run as the comparison baseline."""
    try:
        BASELINE.write_text(json.dumps(_run_payload(res), ensure_ascii=False, indent=2,
                                        default=_json_safe), encoding="utf-8")
        return True
    except Exception:
        return False


def baseline_read() -> dict | None:
    """Scalars needed for the comparison. Deliberately not a full rebuild: the
    baseline is only ever read for its headline figures and its peer set."""
    if not BASELINE.exists():
        return None
    try:
        p = json.loads(BASELINE.read_text(encoding="utf-8"))
    except Exception:
        return None
    rows = p.get("df") or []
    in_stats = [r.get("symbol") for r in rows
                if r.get("included") and not r.get("is_outlier")]
    nd = (p.get("target") or {}).get("net_debt") or 0
    dcf_ev = (p.get("dcf") or {}).get("enterprise_value")
    return {
        "saved_at": p.get("saved_at"),
        "target_name": p.get("target_name"),
        "universe": p.get("universe"),
        "ccy": p.get("ccy"),
        "is_shortlist": p.get("is_shortlist"),
        "n_in_stats": len(in_stats),
        "peers_in_stats": sorted(x for x in in_stats if x),
        "wacc": (p.get("wacc") or {}).get("wacc"),
        "dcf_ev": dcf_ev,
        "dcf_equity": (dcf_ev - nd) if dcf_ev is not None else None,
        "comps_median_ev": (p.get("convergence") or {}).get("comps_median_ev"),
        "multiples": {r["metric"]: r.get("median") for r in (p.get("summary") or [])},
    }


def current_summary(res: dict) -> dict:
    """Same shape as baseline_read(), for the run currently on screen."""
    df = res["df"]
    in_stats = df.loc[df["included"] & ~df["is_outlier"], "symbol"].tolist()
    nd = (res.get("target") or {}).get("net_debt") or 0
    dcf_ev = (res.get("dcf") or {}).get("enterprise_value")
    summary = res["summary"].to_dict("records") if res.get("summary") is not None else []
    return {
        "ccy": res.get("ccy"),
        "is_shortlist": res.get("is_shortlist"),
        "n_in_stats": len(in_stats),
        "peers_in_stats": sorted(in_stats),
        "wacc": (res.get("wacc") or {}).get("wacc"),
        "dcf_ev": dcf_ev,
        "dcf_equity": (dcf_ev - nd) if dcf_ev is not None else None,
        "comps_median_ev": (res.get("convergence") or {}).get("comps_median_ev"),
        "multiples": {r["metric"]: r.get("median") for r in summary},
    }


def snapshot_load() -> dict | None:
    """Rebuild a result bundle from the snapshot. No network, no recomputation."""
    if not SNAPSHOT.exists():
        return None
    try:
        p = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    except Exception:
        return None
    prec = p.get("precedent")
    if prec:
        _pdeals = [Deal(**d) for d in prec["deals"]]
        prec = {"deals": _pdeals,
                "included": prec.get("included") or [True] * len(_pdeals),
                "stats": prec["stats"], "implied": prec["implied"]}
    res = dict(
        df=pd.DataFrame(p["df"]),
        summary=pd.DataFrame(p["summary"]),
        implied=p.get("implied") or [],
        target=p.get("target") or {},
        has_target=p.get("has_target", False),
        wacc=p.get("wacc") or {},
        dcf=p.get("dcf") or {},
        dcf_sens=p.get("dcf_sens") or [],
        convergence=p.get("convergence"),
        precedent=prec,
        log=(p.get("log") or []) + [f"[snapshot] loaded run from {p.get('saved_at')}"],
        sanity=p.get("sanity") or [],
        ccy=p.get("ccy") or "USD",
        n_tickers=p.get("n_tickers"),
        skipped=p.get("skipped") or [],
        is_shortlist=p.get("is_shortlist", False),
        xlsx_path=p.get("xlsx_path"),
        pptx_path=p.get("pptx_path"),
        from_snapshot=p.get("saved_at"),
    )
    # boolean columns survive the JSON round trip as objects in edge cases
    for col in ("included", "is_outlier"):
        if col in res["df"].columns:
            res["df"][col] = res["df"][col].fillna(False).astype(bool)
    return res


def snapshot_stamp() -> str | None:
    if not SNAPSHOT.exists():
        return None
    try:
        return json.loads(SNAPSHOT.read_text(encoding="utf-8")).get("saved_at")
    except Exception:
        return None


load_web_settings()

# Deferred demo-case load: the button in the sidebar sets this flag and reruns, so
# the write happens here, before any widget of those keys exists.
if st.session_state.pop("_load_demo_request", False):
    st.session_state["_demo_loaded"] = load_demo_case()

# Seed universe + tickers on very first load (after settings, before widgets)
if "universe" not in st.session_state:
    st.session_state["universe"] = universe_options()[0]
if "tickers" not in st.session_state:
    try:
        st.session_state["tickers"] = "\n".join(tickers_for(st.session_state["universe"]))
    except Exception:
        st.session_state["tickers"] = ""

# Seed the two-step universe picker from the canonical value
_source_labels = [SOURCE_CURATED] + [AUTO_LABELS[ix] for ix in avail_indices() if ix in AUTO_LABELS]
if "uni_source" not in st.session_state:
    _src, _sub = universe_to_widgets(st.session_state["universe"])
    st.session_state["uni_source"] = _src
    if _src == SOURCE_CURATED:
        st.session_state["uni_curated"] = _sub
    else:
        st.session_state["uni_sector"] = _sub
st.session_state.setdefault("uni_curated", list_universes()[0])
st.session_state.setdefault("uni_sector", "")
# if the stored source is no longer available (e.g. offline), fall back to curated
if st.session_state["uni_source"] not in _source_labels:
    st.session_state["uni_source"] = SOURCE_CURATED


# ---------------------------------------------------------------- CSS + header
_css = APP_DIR / "assets" / "style.css"
if _css.exists():
    st.markdown(f"<style>{_css.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)

ai_available = ai.is_available()
use_ai = bool(st.session_state.get("use_ai", True))
ai_ok = ai_available and use_ai   # AI actually runs only when a key exists AND the toggle is on
_ai_cls = "pg-on" if ai_ok else "pg-off"
_ai_txt = ("AI match on" if ai_ok
           else "AI match off (no key)" if not ai_available
           else "AI match off")
st.markdown(
    f"""
    <div class="pg-mast">
      <div>
        <div class="pg-mast-title">Peer Group Screening</div>
        <div class="pg-mast-sub">Trading comparables, WACC-based DCF, precedent
        transactions and AI peer matching in one run.</div>
      </div>
      <div class="pg-mast-meta">
        <span class="{_ai_cls}">{_ai_txt}</span>
        <span>v{APP_VERSION} web</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------- sidebar
def sb_header(txt: str):
    st.markdown(f"<div class='sb-sec'>{txt}</div>", unsafe_allow_html=True)


def _current_universe() -> str:
    """Canonical universe string from the two-step picker."""
    src = st.session_state.get("uni_source", SOURCE_CURATED)
    if src in LABEL_TO_INDEX:
        return auto_canonical(LABEL_TO_INDEX[src], st.session_state.get("uni_sector", ""))
    return st.session_state.get("uni_curated", list_universes()[0])


def _sync_universe():
    """On source/sector change: recompute universe, refill tickers + target name."""
    # repair the sector when switching to an index with a different taxonomy
    src = st.session_state.get("uni_source", SOURCE_CURATED)
    if src in LABEL_TO_INDEX:
        secs = sectors_for(LABEL_TO_INDEX[src])
        if st.session_state.get("uni_sector") not in secs:
            st.session_state["uni_sector"] = secs[0] if secs else ""
    u = _current_universe()
    st.session_state["universe"] = u
    try:
        st.session_state["tickers"] = "\n".join(tickers_for(u))
    except Exception:
        st.session_state["tickers"] = ""
    st.session_state["target_name"] = clean_target_name(u)


with st.sidebar:
    sb_header("Setup")
    st.text_input("Target name", key="target_name",
                  help="Free text. Becomes the output file name.")

    _sources = [SOURCE_CURATED] + [AUTO_LABELS[ix] for ix in avail_indices() if ix in AUTO_LABELS]
    if len(_sources) > 1:
        st.radio("Universe source", _sources, key="uni_source", horizontal=True,
                 on_change=_sync_universe)
    _src = st.session_state.get("uni_source", SOURCE_CURATED)
    if _src in LABEL_TO_INDEX:
        _idx = LABEL_TO_INDEX[_src]
        _secs = sectors_for(_idx)
        if st.session_state.get("uni_sector") not in _secs:
            st.session_state["uni_sector"] = _secs[0] if _secs else ""
        _tax = "GICS" if _idx == discovery.INDEX_SP500 else "ICB"
        st.selectbox(f"{_tax} sector", _secs, key="uni_sector",
                     on_change=_sync_universe)
        st.caption(AUTO_CAPTION.get(_idx, "") + " Edit the ticker list below to refine.")
    else:
        st.selectbox("Starting universe", list_universes(), key="uni_curated",
                     format_func=pretty_universe, on_change=_sync_universe)
        st.caption("Hand-picked peer set · edit the ticker list below to refine.")
    # keep canonical value consistent on every rerun (not just on change)
    st.session_state["universe"] = _current_universe()

    st.selectbox("Target currency", ["USD", "EUR", "GBP", "CHF"], key="currency",
                 help="All monetary figures are converted to this currency.")
    st.checkbox("Also build PowerPoint deck (.pptx)", key="make_pptx")

    sb_header("Peer universe")
    st.text_area("Tickers (editable, one per line or comma-separated)",
                 key="tickers", height=140,
                 help="Pre-filled from the selected sector. Add or remove tickers freely. "
                      "the analysis always uses THIS list.")
    _n_tickers = len(parse_tickers(st.session_state.get("tickers", "")))
    st.caption(f"{_n_tickers} tickers in the universe")

    sb_header("Screening filters")
    c1, c2 = st.columns(2)
    c1.text_input(f"Revenue min (m {st.session_state['currency']})", key="rev_min")
    c2.text_input(f"Revenue max (m {st.session_state['currency']})", key="rev_max")
    c3, c4 = st.columns(2)
    c3.text_input("EBITDA margin min %", key="mar_min")
    c4.text_input("EBITDA margin max %", key="mar_max")
    st.caption("Empty field = filter off")

    # Everything below is set once per target and then left alone, while the peer
    # set above gets reworked repeatedly. Collapsed by default so the controls in
    # active use stay reachable without scrolling past assumptions.
    _n_target = sum(1 for k in ("t_revenue", "t_ebitda", "t_ebit", "t_net_income",
                                "t_net_debt") if (st.session_state.get(k) or "").strip())
    with st.expander(f"Target financials · {_n_target}/5 filled", expanded=_n_target == 0):
        st.caption(f"in millions {st.session_state['currency']}, needed for implied valuation")
        f1, f2 = st.columns(2)
        f1.text_input("Revenue", key="t_revenue")
        f2.text_input("EBITDA", key="t_ebitda")
        f3, f4 = st.columns(2)
        f3.text_input("EBIT", key="t_ebit")
        f4.text_input("Net income", key="t_net_income")
        st.text_input("Net debt", key="t_net_debt")
        for _w in target_sanity({"revenue": mio(st.session_state["t_revenue"]),
                                 "ebitda": mio(st.session_state["t_ebitda"]),
                                 "ebit": mio(st.session_state["t_ebit"]),
                                 "net_income": mio(st.session_state["t_net_income"])}):
            st.warning(_w)

    _wacc_hint = (f"Rf {st.session_state.get('w_rf') or '-'}% · "
                  f"MRP {st.session_state.get('w_mrp') or '-'}% · "
                  f"D/E {st.session_state.get('w_target_de') or '-'}")
    with st.expander(f"WACC / DCF · {_wacc_hint}", expanded=False):
        w1, w2 = st.columns(2)
        w1.text_input("Risk-free %", key="w_rf")
        w2.text_input("Market premium %", key="w_mrp")
        w3, w4 = st.columns(2)
        w3.text_input("Tax %", key="w_tax")
        w4.text_input("Cost of debt %", key="w_kd")
        w5, w6 = st.columns(2)
        w5.text_input("Target D/E", key="w_target_de")
        w6.text_input("Size premium %", key="w_size_prem")
        d1, d2, d3 = st.columns(3)
        d1.text_input("Base FCF (m)", key="d_base_fcf",
                      help="Leave blank to derive it from the cash-flow bridge below.")
        d2.text_input("FCF growth %", key="d_growth")
        d3.text_input("Terminal growth %", key="d_term_g")
        d4, d5 = st.columns(2)
        d4.text_input("Capex (% of revenue)", key="d_capex_pct")
        d5.text_input("ΔNWC (% of revenue growth)", key="d_nwc_pct")
        st.caption("If Base FCF is blank, the DCF derives it: "
                   "FCF = EBIT·(1−tax) + D&A − Capex − ΔNWC, with D&A = EBITDA − EBIT. "
                   "Needs target EBIT, EBITDA and revenue.")

    with st.expander(f"Precedent transactions · {st.session_state.get('tx_source', '')}",
                     expanded=False):
        st.radio("Deal source", [TX_CSV_LABEL, TX_FMP_LABEL], key="tx_source",
                 horizontal=True)
        if st.session_state["tx_source"] == TX_CSV_LABEL:
            st.text_input("CSV path", key="tx_csv")
            st.file_uploader("…or upload a deal CSV", type=["csv"], key="tx_upload",
                             help="Uploaded file takes precedence over the path above. "
                                  "Columns as in transactions_template.csv.")
        else:
            _k = fmp_key()
            st.caption("FMP key found in config/env" if _k
                       else "No FMP key in config.json or FMP_API_KEY env")
            st.text_input("API key override (optional)", key="tx_key_override",
                          type="password")

    _desc_len = len((st.session_state.get("target_desc") or "").strip())
    _ai_hint = ("on" if ai_ok else "off (no key)" if not ai_available else "off")
    with st.expander(f"AI peer match · {_ai_hint}", expanded=False):
        if ai_available:
            st.toggle("Use AI peer matching", key="use_ai",
                      help="Off = no call to the Gemini API. Nothing about the target "
                           "leaves this machine. Screening, comps, WACC and DCF all still run.")
        else:
            st.caption("AI matching needs GEMINI_API_KEY in .env. Currently off — "
                       "everything else runs without it.")
        st.text_area("Target business model", key="target_desc", height=90,
                     placeholder="What it does, customers, distribution, region, "
                                 "revenue model …")
        if ai_available and not use_ai:
            st.caption("AI is off: the description is not sent anywhere.")

    st.write("")
    run = st.button("Run peer group analysis", type="primary", use_container_width=True)

    sb_header("Demo")
    if DEMO_CASE.exists():
        if st.button("Load demo case", use_container_width=True,
                     help="Restores the curated demo inputs. Does not run the analysis."):
            st.session_state["_load_demo_request"] = True
            st.rerun()
    _stamp = snapshot_stamp()
    if _stamp:
        if st.button("Load last results", use_container_width=True,
                     help="Shows the last completed run from cache. No network calls, "
                          "so this works offline and instantly."):
            _snap = snapshot_load()
            if _snap:
                st.session_state["res"] = _snap
                st.session_state["base_df"] = _snap["df"]
                st.rerun()
        st.caption(f"Cached run from {_stamp.replace('T', ' ')}")
    else:
        st.caption("No cached run yet. Runs are saved automatically.")



# ---------------------------------------------------------------- compute
def target_inputs() -> dict:
    return {"revenue": mio(st.session_state["t_revenue"]),
            "ebitda": mio(st.session_state["t_ebitda"]),
            "ebit": mio(st.session_state["t_ebit"]),
            "net_income": mio(st.session_state["t_net_income"]),
            "net_debt": mio(st.session_state["t_net_debt"])}


def transactions_config() -> tuple[str, str | Path, str]:
    """Resolve deal source → (source_kind, csv_path, api_key)."""
    if st.session_state["tx_source"] == TX_FMP_LABEL:
        key = (st.session_state.get("tx_key_override") or "").strip() or fmp_key()
        return "fmp", "", key
    up = st.session_state.get("tx_upload")
    if up is not None:
        CACHE_DIR.mkdir(exist_ok=True)
        path = CACHE_DIR / "uploaded_transactions.csv"
        path.write_bytes(up.getvalue())
        return "csv", path, ""
    raw = (st.session_state.get("tx_csv") or "").strip()
    if not raw:
        return "csv", "", ""
    path = Path(raw) if Path(raw).is_absolute() else APP_DIR / raw
    return "csv", path, ""


def compute_bundle(df: pd.DataFrame, log: list[str]) -> dict:
    """Takes a filtered+flagged df, returns all valuation results. No fetching."""
    target = target_inputs()
    has_target = any(v is not None for v in target.values())
    sanity = target_sanity(target)
    for w in sanity:
        log.append(f"[check] Target: {w}")
    summary = summary_stats(df[df["included"]], exclude_outliers=True)
    implied = compute_implied(target, summary) if has_target else []
    if implied:
        log.append(f"Implied valuation: {len(implied)} multiples")

    # ---- WACC + DCF ----
    wacc_result, dcf_result, dcf_sens = {}, {}, []
    rf, mrp, tax, kd = (pct(st.session_state["w_rf"]), pct(st.session_state["w_mrp"]),
                        pct(st.session_state["w_tax"]), pct(st.session_state["w_kd"]))
    de_raw = (st.session_state["w_target_de"] or "").strip().replace(",", ".")
    de = float(de_raw) if de_raw else None
    if all(v is not None for v in (rf, mrp, tax, kd)) and de is not None:
        pool = df[df["included"]]
        if "match_score" in df.columns and df["match_score"].notna().any():
            m = pool[pd.to_numeric(pool["match_score"], errors="coerce") >= 6]
            if len(m) >= 3:
                pool = m
        wacc_result = compute_wacc(
            pool[["symbol", "name", "beta", "market_cap", "net_debt"]].to_dict("records"),
            rf, mrp, tax, kd, de, size_premium=pct(st.session_state["w_size_prem"]) or 0.0)
        if wacc_result:
            log.append(f"WACC: {wacc_result['wacc']:.2%} (beta {wacc_result['relevered_beta']:.2f})")
            if wacc_result.get("n_de_outliers"):
                wc = wacc_result.get("wacc_clean")
                log.append(f"[check] {wacc_result['n_de_outliers']} peer(s) with D/E >200%, "
                           f"beta may be distorted{f', cleaned ~{wc:.2%}' if wc else ''}")
        base_fcf = mio(st.session_state["d_base_fcf"])
        growth, term_g = pct(st.session_state["d_growth"]), pct(st.session_state["d_term_g"])
        # Cash-flow bridge: if no Base FCF was typed, derive it from operating metrics
        # (FCFF). D&A comes free from EBITDA − EBIT; Capex and ΔNWC are % inputs.
        bridge = None
        if base_fcf is None and growth is not None and tax is not None:
            ebit = target.get("ebit"); ebitda = target.get("ebitda"); revenue = target.get("revenue")
            if ebit is not None and revenue is not None:
                da = (ebitda - ebit) if (ebitda is not None) else 0.0
                if ebitda is None:
                    # Silent D&A=0 badly understates the FCF (and the DCF value). Surface it
                    # instead of quietly shipping a wrong number.
                    _w = ("No EBITDA entered, so depreciation & amortisation is taken as 0 — "
                          "the derived free cash flow (and the DCF value) is understated. "
                          "Enter EBITDA, or type a Base FCF directly.")
                    sanity.append(_w)
                    log.append(f"[check] {_w}")
                capex_pct = pct(st.session_state["d_capex_pct"]) or 0.0
                nwc_pct = pct(st.session_state["d_nwc_pct"]) or 0.0
                capex = revenue * capex_pct
                delta_nwc = revenue * growth * nwc_pct
                base_fcf = fcf_bridge(ebit, tax, da, capex, delta_nwc)
                bridge = {"ebit": ebit, "tax": tax, "da": da, "revenue": revenue,
                          "growth": growth, "capex_pct": capex_pct, "nwc_pct": nwc_pct,
                          "capex": capex, "delta_nwc": delta_nwc}
                log.append(f"FCF bridge: base FCF {base_fcf/SCALE:,.0f}m "
                           f"(NOPAT {ebit*(1-tax)/SCALE:,.0f} + D&A {da/SCALE:,.0f} "
                           f"− Capex {capex/SCALE:,.0f} − ΔNWC {delta_nwc/SCALE:,.0f})")
        if wacc_result and base_fcf and growth is not None and term_g is not None:
            dcf_result = dcf(base_fcf, growth, term_g, wacc_result["wacc"], bridge=bridge)
            dcf_sens = dcf_sensitivity(base_fcf, growth, term_g, wacc_result["wacc"])
            if dcf_result.get("enterprise_value"):
                log.append(f"DCF EV: {dcf_result['enterprise_value']/SCALE:,.0f} m "
                           f"{st.session_state['currency']}")

    # ---- Convergence check: DCF vs comps median EV ----
    convergence = None
    if dcf_result.get("enterprise_value") and implied:
        cev = [r["base_ev"] for r in implied if r.get("base_ev")]
        if cev:
            cm = sorted(cev)[len(cev) // 2]
            gap = abs(dcf_result["enterprise_value"] - cm) / cm
            convergence = {"comps_median_ev": cm,
                           "dcf_ev": dcf_result["enterprise_value"], "gap": gap}
            if gap > 0.5:
                log.append(f"[check] DCF and comps diverge by {gap:.0%}, review WACC/assumptions")

    # ---- Precedent transactions (4th method) — pluggable source ----
    precedent = None
    try:
        source, csv_path, api_key = transactions_config()
        if has_target and (source == "fmp" or csv_path):
            if source == "csv":
                all_deals = load_transactions(csv_path, scale=SCALE, source="csv")
                deals = load_transactions(csv_path, scale=SCALE, source="csv",
                                          universe_name=st.session_state["universe"])
                if all_deals:
                    log.append(f"Precedent transactions: {len(all_deals)} deals in file, "
                               f"{len(deals)} match the sector")
            else:
                log.append("Deal source: FMP API (live feed) …")
                all_deals = load_transactions(source="fmp", api_key=api_key)
                deals = all_deals  # feed has no sector tag → no filter
                n_valued = sum(1 for d in deals if d.ev)
                log.append(f"FMP: {len(all_deals)} deals loaded, {n_valued} with deal value")
                if all_deals and n_valued == 0:
                    log.append("[check] FMP free tier returns no deal values, so no "
                               "multiples (adapter works, data missing)")
            if deals:
                stats = deal_multiple_stats(deals)
                ti = implied_from_transactions(target, stats)
                if ti:
                    precedent = {"deals": deals, "included": [True] * len(deals),
                                 "stats": stats, "implied": ti}
    except Exception as e:
        log.append(f"Transactions error: {e}")

    return dict(df=df, summary=summary, implied=implied, target=target, has_target=has_target,
                wacc=wacc_result, dcf=dcf_result, dcf_sens=dcf_sens, convergence=convergence,
                precedent=precedent, log=log, sanity=sanity)


def export_run(res: dict) -> None:
    """Write Excel (+ optional PPTX) into a per-run output folder.
    Stores file paths in res — called ONCE per run/shortlist, not on rerun."""
    ccy = res["ccy"]
    df = res["df"]
    name = (st.session_state["target_name"] or "peer_group").strip().replace(" ", "_")
    tag = "_SHORTLIST" if res.get("is_shortlist") else ""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = OUTPUT_DIR / f"{name}{tag}_{ts}"
    run_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = run_dir / f"{name}{tag}.xlsx"

    meta = {"Target": name + ("  [SHORTLIST]" if tag else ""),
            "Sector / universe": st.session_state["universe"],
            "Target currency": ccy,
            "Tickers evaluated": res.get("n_tickers", len(df)),
            "Data received": len(df),
            "Skipped (no data)": len(res.get("skipped") or []),
            "Included (pass filters)": int(df["included"].sum()),
            "Outliers flagged": int(df["is_outlier"].sum()),
            f"Revenue range (m {ccy})":
                f"{st.session_state['rev_min'] or '–'} – {st.session_state['rev_max'] or '–'}",
            "EBITDA margin range":
                f"{st.session_state['mar_min'] or '–'}% – {st.session_state['mar_max'] or '–'}%",
            "Run timestamp": datetime.now().isoformat(timespec="seconds")}
    if tag:
        meta["Mode"] = "Short list (manually selected peers)"

    export(df, res["summary"], meta, xlsx_path, skipped=res.get("skipped"),
           implied=res["implied"], target_inputs=res["target"] if res["has_target"] else None,
           wacc_result=res["wacc"] or None, dcf_result=res["dcf"] or None,
           dcf_sens=res["dcf_sens"] or None, convergence=res["convergence"],
           precedent=res["precedent"])
    res["xlsx_path"] = str(xlsx_path)
    res["log"].append(f"Excel written: {xlsx_path}")

    res["pptx_path"], res["pptx_error"] = None, None
    n_incl = int((df["included"] & ~df["is_outlier"]).sum())
    if st.session_state["make_pptx"] and res["has_target"]:
        try:
            from pptx_export import build_pptx
            pptx_path = xlsx_path.with_suffix(".pptx")
            build_pptx(pptx_path, target_name=name,
                       sector=st.session_state["universe"], ccy=ccy,
                       summary_df=res["summary"], implied=res["implied"],
                       wacc_result=res["wacc"] or None, dcf_result=res["dcf"] or None,
                       precedent=res["precedent"], n_peers_included=n_incl)
            res["pptx_path"] = str(pptx_path)
            res["log"].append(f"PowerPoint written: {pptx_path}")
        except Exception as e:
            res["pptx_error"] = str(e)
            res["log"].append(f"PowerPoint error: {e}")
    res["log"].append(f"Run folder: {run_dir}")


def run_analysis():
    tickers = parse_tickers(st.session_state.get("tickers", ""))
    if not tickers:
        st.error("No tickers in the universe. Pick a sector or add tickers.")
        return None
    save_web_settings()  # persist this search, like the desktop app
    ccy = st.session_state["currency"]
    log: list[str] = [f"[1/4] Universe: {len(tickers)} tickers · target currency {ccy}",
                      f"      Filters: revenue {st.session_state['rev_min'] or '–'}–"
                      f"{st.session_state['rev_max'] or '–'} m {ccy}"]
    ds = YFinanceDataSource(target_currency=ccy)
    rows, skipped = [], []

    with st.status(f"Fetching financials for {len(tickers)} tickers …", expanded=True) as status:
        prog = st.progress(0.0)
        log.append(f"[2/4] Fetching financials from Yahoo Finance "
                   f"({FETCH_WORKERS} parallel requests) …")
        # Fetch in a thread pool: the run time is dominated by network latency, not
        # by CPU. Only ds.fetch() runs in the workers — every Streamlit call
        # (progress, status) stays on the main thread, because the Streamlit script
        # context does not exist in worker threads.
        results: dict[int, object] = {}
        done = 0
        with ThreadPoolExecutor(max_workers=FETCH_WORKERS) as pool:
            futures = {pool.submit(ds.fetch, sym): (i, sym)
                       for i, sym in enumerate(tickers, 1)}
            for fut in as_completed(futures):
                i, sym = futures[fut]
                done += 1
                prog.progress(done / len(tickers), text=f"[{done}/{len(tickers)}] {sym}")
                try:
                    results[i] = fut.result()
                except Exception as e:
                    results[i] = e

        # Replay in ticker order so the log and the resulting frame stay
        # deterministic, independent of which request finished first.
        for i, sym in enumerate(tickers, 1):
            fin = results.get(i)
            if isinstance(fin, Exception):
                skipped.append((sym, f"error: {fin}"))
                log.append(f"   [{i:>3}/{len(tickers)}] {sym}: error: {fin}")
                continue
            if fin is None:
                skipped.append((sym, "no equity data / wrong quote type"))
                log.append(f"   [{i:>3}/{len(tickers)}] {sym}: skipped (no equity data)")
                continue
            rows.append(fin.__dict__)
            rev_str = f"{fin.revenue/SCALE:>7.0f}M" if fin.revenue else "    n/a"
            log.append(f"   [{i:>3}/{len(tickers)}] {sym:<12} "
                       f"{(fin.name or '')[:30]:<30} rev={rev_str}")

        if not rows:
            status.update(label="No data received.", state="error")
            st.error("No data received. Check the tickers and the connection.")
            return None
        df = pd.DataFrame(rows)

        # ---- AI business-model match ----
        # Gated on the toggle: with it off, no request is made to Gemini, so nothing
        # about the target leaves the machine. Logged either way for the audit trail.
        target_desc = (st.session_state["target_desc"] or "").strip()
        _use_ai = bool(st.session_state.get("use_ai", True))
        if target_desc and ai.is_available() and not _use_ai:
            log.append("[AI] off by choice — no data sent to Gemini; matching skipped.")
        if target_desc and ai.is_available() and _use_ai:
            status.update(label="AI is scoring business-model fit …")
            log.append(f"      AI scoring business-model match (Gemini, 1 call for {len(df)} peers) …")
            try:
                sc = ai.score_peers(target_desc,
                                    df[["symbol", "name", "sector", "industry",
                                        "business_summary"]].to_dict("records"))
                if sc:
                    df["match_score"] = df["symbol"].map(lambda s: sc.get(s, {}).get("score"))
                    df["match_rationale"] = df["symbol"].map(lambda s: sc.get(s, {}).get("rationale"))
                    df["business_profile"] = df["symbol"].map(lambda s: sc.get(s, {}).get("business_profile"))
                    log.append(f"      AI match scores: "
                               f"{sum(1 for v in sc.values() if v.get('score'))} peers scored")
                else:
                    log.append("      AI returned no usable answer.")
            except Exception as e:
                log.append(f"      AI error: {e}")
                st.warning(f"AI error: {e}")

        status.update(label=f"Fetched {len(rows)}/{len(tickers)} tickers, computing",
                      state="running")
        log.append("[3/4] Filters + outlier detection …")
        df = apply_filters(df, mio(st.session_state["rev_min"]), mio(st.session_state["rev_max"]),
                           pct(st.session_state["mar_min"]), pct(st.session_state["mar_max"]))
        df = flag_outliers(df)
        if "match_score" in df.columns and df["match_score"].notna().any():
            df["_s"] = df["match_score"].fillna(-1)
            df = (df.sort_values(["included", "_s"], ascending=[False, False])
                    .drop(columns="_s").reset_index(drop=True))
        else:
            rm = ((mio(st.session_state["rev_min"]) or 0) +
                  (mio(st.session_state["rev_max"]) or 0)) / 2 or None
            df = rank_by_similarity(df, target_revenue=rm)

        log.append("[4/4] Valuation + export …")
        res = compute_bundle(df, log)
        res.update(ccy=ccy, n_tickers=len(tickers), skipped=skipped, is_shortlist=False)
        export_run(res)
        status.update(label=f"Done. {len(rows)}/{len(tickers)} tickers analysed.",
                      state="complete", expanded=False)
    return res


if run:
    r = run_analysis()
    if r:
        st.session_state["res"] = r
        st.session_state["base_df"] = r["df"]
        snapshot_save(r)

res = st.session_state.get("res")

if st.session_state.pop("_demo_loaded", False):
    st.success("Demo case loaded. Hit Run peer group analysis, or Load last results "
               "to show the cached run without any network calls.")

# ---------------------------------------------------------------- empty state
if not res:
    _methods = [
        ("Trading comparables",
         "Live peer multiples (EV/Sales, EV/EBITDA, EV/EBIT, P/E) with IQR outlier "
         "flagging and quartile statistics."),
        ("DCF on a peer-derived WACC",
         "Peer betas unlevered and relevered to the target capital structure, then a "
         "5-year FCF projection with terminal value and a sensitivity grid."),
        ("Precedent transactions",
         "Deal multiples from your own CSV or a live FMP feed, read as the control-premium "
         "anchor above trading comps."),
        ("AI peer match",
         "Gemini scores each peer's business-model fit to the target (1 to 10) with a "
         "written rationale. It ranks comparability and never touches the figures."),
    ]
    _rows = "".join(
        f"<div class='pg-method'><div class='pg-method-name'>{n}</div>"
        f"<div class='pg-method-body'>{t}</div></div>" for n, t in _methods)
    st.markdown(
        "<div class='pg-lede'>Set the parameters in the sidebar, then run the analysis. "
        "One run produces four valuation angles plus a formatted Excel workbook and an "
        "optional PowerPoint deck.</div>"
        f"<div class='pg-methods'>{_rows}</div>",
        unsafe_allow_html=True)
    st.markdown(
        f"<div class='pg-footer'>Peer Group Screening <b>v{APP_VERSION}</b> · "
        "Data: Yahoo Finance, FMP · For discussion purposes only. Not investment advice."
        "</div>", unsafe_allow_html=True)
    st.stop()

# ---------------------------------------------------------------- results
df = res["df"]
ccy = res["ccy"]
n_incl = int((df["included"] & ~df["is_outlier"]).sum())

if res.get("from_snapshot"):
    st.info(f"Cached run from {res['from_snapshot'].replace('T', ' ')}. No data was "
            f"fetched for this view. Run the analysis again for live figures.")

for _w in res.get("sanity") or []:
    st.warning(f"Target check: {_w}")

if res.get("is_shortlist"):
    st.success("Short list active. Statistics are based on the selected peers only.")

if res.get("manual_override"):
    st.success("Manual peer selection active. The automatic filters and the IQR rule "
               "made a proposal, the peer set below is the one you decided on.")


def kpi_row(items: list[tuple[str, str]]):
    """Hairline-divided figure strip. No card boxes: at this density the rules
    carry the structure and the numbers stay the loudest thing on screen."""
    cells = "".join(
        f"<div class='pg-strip-item'><div class='pg-strip-label'>{lbl}</div>"
        f"<div class='pg-strip-value'>{val}</div></div>" for lbl, val in items)
    st.markdown(f"<div class='pg-strip'>{cells}</div>", unsafe_allow_html=True)


def funnel_row(steps: list[tuple[int, str]]):
    """How the peer group narrowed, in one line. This is the answer to the first
    question any reviewer asks: how did you arrive at this peer set?"""
    cells = []
    for i, (n, lbl) in enumerate(steps):
        if i:
            cells.append("<span class='pg-funnel-arrow'>&rarr;</span>")
        cells.append(f"<span class='pg-funnel-step'><span class='pg-funnel-n'>{n}</span>"
                     f"<span class='pg-funnel-l'>{lbl}</span></span>")
    st.markdown(f"<div class='pg-funnel'>{''.join(cells)}</div>", unsafe_allow_html=True)


_n_fetched = int(res.get("n_tickers") or len(df))
_n_data = len(df)
_n_filtered = int(df["included"].sum())
funnel_row([
    (_n_fetched, "in universe"),
    (_n_data, "with data"),
    (_n_filtered, "pass filters"),
    (n_incl, "in statistics"),
])

_cm = res["convergence"]["comps_median_ev"] / SCALE if res["convergence"] else None
_gap = res["convergence"]["gap"] if res["convergence"] else None
kpi_row([
    ("WACC", f"{res['wacc']['wacc']:.2%}" if res["wacc"] else "n/a"),
    (f"DCF EV (m {ccy})",
     f"{res['dcf']['enterprise_value']/SCALE:,.0f}"
     if res["dcf"].get("enterprise_value") else "n/a"),
    (f"Comps median EV (m {ccy})", f"{_cm:,.0f}" if _cm else "n/a"),
    ("DCF vs comps gap", f"{_gap:.0%}" if _gap is not None else "n/a"),
])

# ---------------------------------------------------------------- compare
# Controls live here rather than in the sidebar: the sidebar is rendered before
# the run result exists, so a "pin this run" button there would only appear one
# interaction late.
_base = baseline_read()
_pin1, _pin2 = st.columns([1, 3])
if _pin1.button("Pin as baseline", use_container_width=True,
                help="Freezes this run. Every later run is then shown with the "
                     "difference against it, including which peers moved."):
    if baseline_save(res):
        st.rerun()
if _base and _pin2.button("Clear baseline", use_container_width=False):
    try:
        BASELINE.unlink()
    except OSError:
        pass
    st.rerun()
if not _base:
    _pin2.caption("No baseline pinned yet. Pin a run, change the peer set or an "
                  "assumption, and the next run shows the difference.")

if _base:
    _cur = current_summary(res)
    with st.expander("Compare against pinned baseline", expanded=True):
        if _base.get("ccy") != _cur.get("ccy"):
            st.warning(f"Baseline is in {_base.get('ccy')}, this run in {_cur.get('ccy')}. "
                       f"The money figures are not comparable.")

        def _delta_row(label, base, cur, fmt, pct=True):
            if base is None or cur is None:
                return {"Figure": label, "Baseline": "n/a", "Current": "n/a",
                        "Change": "n/a"}
            d = cur - base
            rel = f"{d / base:+.1%}" if (pct and base) else ""
            return {"Figure": label, "Baseline": fmt(base), "Current": fmt(cur),
                    "Change": f"{fmt(d) if d < 0 else '+' + fmt(d)}"
                              + (f"  ({rel})" if rel else "")}

        _m = lambda v: f"{v / SCALE:,.0f}"
        _p = lambda v: f"{v:.2%}"
        _n = lambda v: f"{v:,.0f}"
        rows_cmp = [
            _delta_row("Peers in statistics", _base["n_in_stats"], _cur["n_in_stats"], _n),
            _delta_row("WACC", _base["wacc"], _cur["wacc"], _p),
            _delta_row(f"DCF enterprise value (m {ccy})", _base["dcf_ev"], _cur["dcf_ev"], _m),
            _delta_row(f"DCF equity value (m {ccy})", _base["dcf_equity"],
                       _cur["dcf_equity"], _m),
            _delta_row(f"Comps median EV (m {ccy})", _base["comps_median_ev"],
                       _cur["comps_median_ev"], _m),
        ]
        for _key, _lbl in [("ev_sales", "EV/Sales median"),
                           ("ev_ebitda", "EV/EBITDA median"),
                           ("ev_ebit", "EV/EBIT median"),
                           ("pe", "P/E median")]:
            rows_cmp.append(_delta_row(_lbl, _base["multiples"].get(_key),
                                       _cur["multiples"].get(_key),
                                       lambda v: f"{v:.2f}x"))
        st.dataframe(pd.DataFrame(rows_cmp), use_container_width=True, hide_index=True)

        # The peer set is the actual analyst judgement, so name the difference
        # explicitly rather than leaving it to be read off two tables.
        _added = sorted(set(_cur["peers_in_stats"]) - set(_base["peers_in_stats"]))
        _removed = sorted(set(_base["peers_in_stats"]) - set(_cur["peers_in_stats"]))
        if _added or _removed:
            st.markdown("##### Peer set")
            if _added:
                st.markdown(f"**Added:** {', '.join(_added)}")
            if _removed:
                st.markdown(f"**Removed:** {', '.join(_removed)}")
        else:
            st.caption("Same peers in the statistics as the baseline. Any change above "
                       "comes from the assumptions, not from the peer set.")
        st.caption(f"Baseline pinned {(_base.get('saved_at') or '').replace('T', ' ')}"
                   + (f" · {_base['target_name']}" if _base.get("target_name") else "")
                   + (" · short list" if _base.get("is_shortlist") else ""))

with st.expander(f"Run log · {len(res.get('skipped') or [])} tickers skipped"):
    st.code("\n".join(res.get("log") or ["(no log)"]), language=None)
    if res.get("skipped"):
        st.dataframe(pd.DataFrame(res["skipped"], columns=["Ticker", "Reason"]),
                     use_container_width=True, hide_index=True)

# Order follows the analyst's sequence: see the peers before selecting them.
# Football Field stays first because the answer belongs at the top when presenting.
tabs = st.tabs(["Football Field", "Long List", "Short List", "Summary",
                "Valuation", "WACC / DCF", "Precedent Tx", "AI Match",
                "Methodology"])

# --- Football Field ---
with tabs[0]:
    methods = []
    for it in (res["implied"] or []):
        if it.get("base_ev"):
            methods.append((it["label"], (it.get("low_ev") or it["base_ev"]) / SCALE,
                            it["base_ev"] / SCALE, (it.get("high_ev") or it["base_ev"]) / SCALE))
    if res["dcf"].get("enterprise_value"):
        ev = res["dcf"]["enterprise_value"] / SCALE
        methods.append(("DCF", ev * 0.9, ev, ev * 1.1))
    if res["precedent"]:
        pk = next((t for t in res["precedent"]["implied"] if t["multiple"] == "ev_ebitda"), None)
        if pk and pk.get("base_ev"):
            methods.append(("Precedent Tx", (pk.get("low_ev") or pk["base_ev"]) / SCALE,
                            pk["base_ev"] / SCALE, (pk.get("high_ev") or pk["base_ev"]) / SCALE))
    if methods:
        labels = [m[0] for m in methods]
        lows = [m[1] for m in methods]
        meds = [m[2] for m in methods]
        highs = [m[3] for m in methods]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=labels, x=[h - l for l, h in zip(lows, highs)], base=lows, orientation="h",
            marker=dict(color="#C6D8D2", line=dict(color=ACCENT, width=1)),
            text=[f"{l:,.0f} to {h:,.0f}" for l, h in zip(lows, highs)],
            textposition="inside", insidetextanchor="middle",
            textfont=dict(color=INK, size=11),
            customdata=list(zip(lows, meds, highs)),
            hovertemplate="<b>%{y}</b><br>Low %{customdata[0]:,.0f}<br>"
                          "Median %{customdata[1]:,.0f}<br>High %{customdata[2]:,.0f}"
                          "<extra></extra>",
            showlegend=False))
        fig.add_trace(go.Scatter(
            y=labels, x=meds, mode="markers", name="Median",
            marker=dict(symbol="line-ns", size=26, line=dict(color=INK, width=3)),
            hovertemplate="Median %{x:,.0f}<extra></extra>", showlegend=False))
        if res["dcf"].get("enterprise_value"):
            fig.add_vline(x=res["dcf"]["enterprise_value"] / SCALE, line_dash="dot",
                          line_color=ACCENT, line_width=1.5,
                          annotation_text="DCF EV", annotation_position="top",
                          annotation_font=dict(size=10, color=ACCENT))
        style_fig(fig, height=104 + 58 * len(methods),
                  title=f"Implied enterprise value range (m {ccy})",
                  xaxis_title=f"m {ccy}", bargap=0.5)
        st.plotly_chart(fig, use_container_width=True)
        if res["convergence"]:
            g = res["convergence"]["gap"]
            (st.success if g <= 0.5 else st.warning)(
                f"Convergence: DCF {res['convergence']['dcf_ev']/SCALE:,.0f} vs "
                f"comps median {res['convergence']['comps_median_ev']/SCALE:,.0f} m {ccy}. "
                f"Gap {g:.0%}.")
        # Every bar is an EV. The bridge to equity is the figure a buyer pays, so it
        # is carried in the same table instead of being left to the export.
        _nd = (res["target"].get("net_debt") or 0) / SCALE
        st.dataframe(pd.DataFrame([{"Method": m[0], "Low": round(m[1]), "Median": round(m[2]),
                                    "High": round(m[3]),
                                    "Median equity": round(m[2] - _nd)} for m in methods]),
                     use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(format="localized")
                                    for c in ("Low", "Median", "High", "Median equity")})
        _evs = [m[2] for m in methods]
        st.caption(f"Bars are enterprise value. Across all methods the implied equity "
                   f"range is {min(_evs) - _nd:,.0f} to {max(_evs) - _nd:,.0f} m {ccy} "
                   f"(EV minus net debt of {_nd:,.0f}). Equity value is what a buyer pays.")
    else:
        st.info("Enter target financials to unlock the valuation methods.")

# --- Short List ---
with tabs[2]:
    st.write("**Select the peers that feed statistics and valuation.** "
             "Pre-selected: match score 7 or higher, otherwise everything that passes "
             "the filters.")
    base_df = st.session_state.get("base_df", df)
    has_scores = "match_score" in base_df.columns and base_df["match_score"].notna().any()

    def _label(r):
        sc = r.get("match_score")
        sc_s = f"[{int(sc)}]" if pd.notna(sc) else "[-]"
        flag = "" if r.get("included") and not r.get("is_outlier") else " (filtered/outlier)"
        return f"{r['symbol']} {sc_s} {(r.get('name') or '')[:30]}{flag}"

    options = {_label(r): r["symbol"] for _, r in base_df.iterrows()}
    if has_scores:
        default = [lbl for lbl, sym in options.items()
                   if pd.notna(base_df.loc[base_df.symbol == sym, "match_score"].iloc[0])
                   and int(base_df.loc[base_df.symbol == sym, "match_score"].iloc[0]) >= 7]
    else:
        default = [lbl for lbl, sym in options.items()
                   if bool(base_df.loc[base_df.symbol == sym, "included"].iloc[0])
                   and not bool(base_df.loc[base_df.symbol == sym, "is_outlier"].iloc[0])]
    picked = st.multiselect("Peers", list(options.keys()), default=default)
    if st.button("Compute short list", type="primary"):
        syms = [options[p] for p in picked]
        if len(syms) < 2:
            st.warning("Pick at least 2 peers.")
        else:
            sub = base_df[base_df["symbol"].isin(syms)].copy()
            sub["included"] = True
            sub["is_outlier"] = False
            sub["excluded_reason"] = ""
            sub["outlier_flags"] = ""
            log = [f"[Short list] {len(syms)} peers selected, recomputing"]
            new = compute_bundle(sub, log)
            new.update(ccy=st.session_state["currency"], n_tickers=len(syms),
                       skipped=res.get("skipped", []), is_shortlist=True)
            export_run(new)
            st.session_state["res"] = new
            snapshot_save(new)
            st.rerun()

# --- Long List ---
with tabs[1]:
    st.write("**The filters and the IQR rule propose a peer set. The tick box decides.** "
             "Untick a peer the statistics should ignore, tick one the rule threw out but "
             "you consider comparable, then recompute.")
    cols = [c for c in ["symbol", "name", "match_score", "country", "revenue", "ebitda",
                        "ebitda_margin", "ev_sales", "ev_ebitda", "ev_ebit", "pe",
                        "excluded_reason", "outlier_flags"] if c in df.columns]
    disp = df[cols].copy()
    for c in ("revenue", "ebitda"):
        if c in disp.columns:
            disp[c] = pd.to_numeric(disp[c], errors="coerce") / SCALE
    # The proposal of the automatic rules, as a column the analyst can overrule.
    # excluded_reason and outlier_flags stay visible so the reason for every
    # proposal is readable next to the decision.
    _proposed = (df["included"] & ~df["is_outlier"]).values
    disp.insert(0, "in_stats", _proposed)
    col_cfg = {
        "in_stats": st.column_config.CheckboxColumn(
            "In stats", help="Feeds the summary statistics, the multiples and the WACC."),
        "symbol": st.column_config.TextColumn("Ticker"),
        "name": st.column_config.TextColumn("Company", width="medium"),
        "match_score": st.column_config.ProgressColumn("AI fit", min_value=0, max_value=10,
                                                       format="%d"),
        "country": st.column_config.TextColumn("Country"),
        "revenue": st.column_config.NumberColumn(f"Revenue (m {ccy})", format="localized"),
        "ebitda": st.column_config.NumberColumn(f"EBITDA (m {ccy})", format="localized"),
        "ebitda_margin": st.column_config.NumberColumn("EBITDA margin", format="percent"),
        "ev_sales": st.column_config.NumberColumn("EV/Sales", format="%.1fx"),
        "ev_ebitda": st.column_config.NumberColumn("EV/EBITDA", format="%.1fx"),
        "ev_ebit": st.column_config.NumberColumn("EV/EBIT", format="%.1fx"),
        "pe": st.column_config.NumberColumn("P/E", format="%.1fx"),
        "excluded_reason": st.column_config.TextColumn("Excluded because"),
        "outlier_flags": st.column_config.TextColumn("Outlier flags"),
    }
    edited = st.data_editor(
        disp, use_container_width=True, hide_index=True, column_config=col_cfg,
        disabled=[c for c in disp.columns if c != "in_stats"],
        key="longlist_editor")

    _chosen = set(edited.loc[edited["in_stats"], "symbol"])
    _auto = set(df.loc[_proposed, "symbol"])
    _added, _dropped = sorted(_chosen - _auto), sorted(_auto - _chosen)
    if _added or _dropped:
        _parts = []
        if _added:
            _parts.append(f"**back in:** {', '.join(_added)}")
        if _dropped:
            _parts.append(f"**taken out:** {', '.join(_dropped)}")
        st.info("Override against the automatic proposal, " + " · ".join(_parts) + ".")
        if len(_chosen) < 2:
            st.warning("At least 2 peers are needed for quartile statistics.")
        elif st.button("Recompute with this selection", type="primary"):
            sub = df.copy()
            sub["included"] = sub["symbol"].isin(_chosen)
            sub["is_outlier"] = False          # the decision is manual from here on
            sub.loc[~sub["included"], "excluded_reason"] = "excluded by the analyst"
            sub.loc[sub["included"], "excluded_reason"] = ""
            log = [f"[Manual override] {len(_chosen)} peers in the statistics "
                   f"(rules proposed {len(_auto)})"]
            if _added:
                log.append(f"      overruled back in: {', '.join(_added)}")
            if _dropped:
                log.append(f"      overruled out: {', '.join(_dropped)}")
            new = compute_bundle(sub, log)
            new.update(ccy=st.session_state["currency"],
                       n_tickers=res.get("n_tickers", len(sub)),
                       skipped=res.get("skipped", []),
                       is_shortlist=res.get("is_shortlist", False),
                       manual_override=True)
            export_run(new)
            st.session_state["res"] = new
            snapshot_save(new)
            st.rerun()
    else:
        st.caption("Ticks match the automatic proposal: everything that passed the "
                   "filters and was not flagged as an outlier.")

# --- Summary ---
with tabs[3]:
    st.dataframe(res["summary"], use_container_width=True, hide_index=True)
    st.caption("Median / mean / quartiles across the included peer set, outliers excluded.")

# --- Valuation ---
with tabs[4]:
    if res["implied"]:
        _nd = (res["target"].get("net_debt") or 0) / SCALE
        st.markdown("##### Implied enterprise value")
        st.dataframe(pd.DataFrame([{"Multiple": r["label"], "Median": round(r["median"], 2),
                                    "Low EV": round((r["low_ev"] or 0) / SCALE),
                                    "Base EV": round(r["base_ev"] / SCALE),
                                    "High EV": round((r["high_ev"] or 0) / SCALE)}
                                   for r in res["implied"]]),
                     use_container_width=True, hide_index=True,
                     column_config={
                         "Median": st.column_config.NumberColumn(format="%.2fx"),
                         "Low EV": st.column_config.NumberColumn(f"Low EV (m {ccy})",
                                                                 format="localized"),
                         "Base EV": st.column_config.NumberColumn(f"Base EV (m {ccy})",
                                                                  format="localized"),
                         "High EV": st.column_config.NumberColumn(f"High EV (m {ccy})",
                                                                  format="localized")})
        st.caption("Implied enterprise value = target metric × peer multiple "
                   "(low = Q1, base = median, high = Q3).")

        # EV is an intermediate figure. What gets negotiated is equity value, so the
        # bridge is shown next to it rather than left in the export.
        st.markdown("##### Implied equity value")
        st.dataframe(pd.DataFrame([{"Multiple": r["label"],
                                    "Low equity": round((r["low_equity"] or 0) / SCALE),
                                    "Base equity": round(r["base_equity"] / SCALE),
                                    "High equity": round((r["high_equity"] or 0) / SCALE)}
                                   for r in res["implied"]]),
                     use_container_width=True, hide_index=True,
                     column_config={
                         "Low equity": st.column_config.NumberColumn(f"Low (m {ccy})",
                                                                     format="localized"),
                         "Base equity": st.column_config.NumberColumn(f"Base (m {ccy})",
                                                                      format="localized"),
                         "High equity": st.column_config.NumberColumn(f"High (m {ccy})",
                                                                      format="localized")})
        st.caption(f"EV-based multiples: equity = EV − net debt ({_nd:,.0f} m {ccy}). "
                   f"P/E works the other way round: it yields equity directly, and EV is "
                   f"derived by adding net debt back.")
    else:
        st.info("No target financials entered.")

# --- WACC / DCF ---
with tabs[5]:
    if res["wacc"]:
        w = res["wacc"]
        kpi_row([("WACC", f"{w['wacc']:.2%}"),
                 ("Relevered beta", f"{w['relevered_beta']:.2f}"),
                 ("Cost of equity", f"{w['cost_of_equity']:.2%}")])
        if w.get("n_de_outliers"):
            st.warning(f"{w['n_de_outliers']} peer(s) with D/E above 200%, so beta may be "
                       f"distorted; cleaned WACC ~{w.get('wacc_clean', 0):.2%}")
        st.dataframe(pd.DataFrame(w["peer_betas"])[["symbol", "beta_levered", "de_ratio",
                                                    "beta_unlevered"]],
                     use_container_width=True, hide_index=True,
                     column_config={
                         "symbol": st.column_config.TextColumn("Ticker"),
                         "beta_levered": st.column_config.NumberColumn("Levered β", format="%.2f"),
                         "de_ratio": st.column_config.NumberColumn("D/E", format="%.2f"),
                         "beta_unlevered": st.column_config.NumberColumn("Unlevered β",
                                                                         format="%.2f")})
    else:
        st.info("Fill all WACC assumptions (risk-free, market premium, tax, cost of debt, D/E).")

    if res["dcf"].get("enterprise_value"):
        d = res["dcf"]
        _nd = (res["target"].get("net_debt") or 0) / SCALE
        _ev = d["enterprise_value"] / SCALE

        # Cash-flow bridge: when the Base FCF was derived (not typed), show where it
        # comes from. Otherwise the derived number is invisible in the UI.
        _br = d.get("bridge")
        if _br:
            st.markdown("##### Cash-flow bridge — derived free cash flow (FCFF)")
            _nopat = _br["ebit"] * (1 - _br["tax"])
            _bridge_df = pd.DataFrame([
                {"Component": "NOPAT = EBIT × (1 − tax)", "m " + ccy: _nopat / SCALE},
                {"Component": "+ D&A  (= EBITDA − EBIT)", "m " + ccy: _br["da"] / SCALE},
                {"Component": f"− Capex  ({_br['capex_pct']:.0%} of revenue)",
                 "m " + ccy: -_br["capex"] / SCALE},
                {"Component": f"− ΔNWC  ({_br['nwc_pct']:.0%} of revenue growth)",
                 "m " + ccy: -_br["delta_nwc"] / SCALE},
                {"Component": "= Base FCF (year 0)", "m " + ccy: (d.get("base_fcf") or 0) / SCALE},
            ])
            st.dataframe(_bridge_df, use_container_width=True, hide_index=True,
                         column_config={"m " + ccy: st.column_config.NumberColumn(
                             f"m {ccy}", format="localized")})
            st.caption("Free cash flow is derived, not entered: EBIT after tax, plus "
                       "non-cash D&A, minus Capex and working-capital investment. This "
                       "Base FCF is what the projection below grows.")

        kpi_row([(f"DCF enterprise value (m {ccy})", f"{_ev:,.0f}"),
                 (f"Net debt (m {ccy})", f"{_nd:,.0f}"),
                 (f"DCF equity value (m {ccy})", f"{_ev - _nd:,.0f}")])
        st.write(f"**DCF EV {d['enterprise_value']/SCALE:,.0f} m {ccy}.** "
                 f"Terminal value is {d['tv_pct_of_ev']:.0%} of EV.")

    if res["dcf_sens"]:
        st.markdown("##### DCF sensitivity: EV by WACC and terminal growth")
        sdf = pd.DataFrame(res["dcf_sens"])
        piv = (sdf.pivot(index="wacc", columns="terminal_growth", values="enterprise_value")
               / SCALE)
        piv = piv.sort_index()
        heat = go.Figure(go.Heatmap(
            z=piv.values,
            x=[f"{c:.2%}" for c in piv.columns],
            y=[f"{i:.2%}" for i in piv.index],
            colorscale=[[0, "#FFFFFF"], [0.5, "#B7CCC5"], [1, ACCENT]],
            text=[[f"{v:,.0f}" for v in row] for row in piv.values],
            texttemplate="%{text}",
            textfont=dict(family="IBM Plex Mono, Consolas, monospace", size=11),
            xgap=1, ygap=1,
            hovertemplate="WACC %{y}, g %{x}<br>EV %{text} m<extra></extra>",
            showscale=False))
        style_fig(heat, height=290, xaxis_title="Terminal growth", yaxis_title="WACC",
                  yaxis_autorange="reversed", margin=dict(l=8, r=14, t=8, b=8))
        heat.update_xaxes(showgrid=False)
        st.plotly_chart(heat, use_container_width=True)

# --- Precedent Tx ---
with tabs[6]:
    if res["precedent"]:
        _deals = res["precedent"]["deals"]
        # All deals stay in the list; only the checkbox toggles. Unselected deals are
        # flagged, not dropped — so you can always tick one back in (like the outliers
        # in the Long List). The applied selection is stored in res["precedent"]["included"].
        _applied = res["precedent"].get("included") or [True] * len(_deals)
        st.write("**Select the deals that feed the precedent statistics.** "
                 "Curate by comparability and recency — a stale or distressed deal is "
                 "not a today's-market price. Unticked deals stay in the list, just out "
                 "of the median.")
        _deal_df = pd.DataFrame([{"include": bool(_applied[i]),
                                  "Date": dd.date, "Target": dd.target,
                                  "Acquirer": dd.acquirer,
                                  "EV/Sales": round(dd.ev_sales, 1) if dd.ev_sales else None,
                                  "EV/EBITDA": round(dd.ev_ebitda, 1) if dd.ev_ebitda else None,
                                  "EV/EBIT": round(dd.ev_ebit, 1) if dd.ev_ebit else None}
                                 for i, dd in enumerate(_deals)])
        _pv = st.session_state.get("precedent_ver", 0)
        _edited_deals = st.data_editor(
            _deal_df, use_container_width=True, hide_index=True,
            key=f"precedent_editor_{_pv}",
            disabled=[c for c in _deal_df.columns if c != "include"],
            column_config={
                "include": st.column_config.CheckboxColumn(
                    "Include", help="Feeds the deal-multiple median and the implied value."),
                "EV/Sales": st.column_config.NumberColumn(format="%.1fx"),
                "EV/EBITDA": st.column_config.NumberColumn(format="%.1fx"),
                "EV/EBIT": st.column_config.NumberColumn(format="%.1fx")})
        _mask = [bool(x) for x in _edited_deals["include"].tolist()]
        _keep = [d for d, inc in zip(_deals, _mask) if inc]
        if _mask != [bool(x) for x in _applied]:
            if len(_keep) < 1:
                st.warning("Keep at least one deal.")
            elif st.button("Recompute precedent with this selection", type="primary"):
                _stats = deal_multiple_stats(_keep)
                _ti = implied_from_transactions(res["target"], _stats)
                _new = dict(res)
                # keep the FULL deal list, store only the mask + recomputed stats
                _new["precedent"] = {"deals": _deals, "included": _mask,
                                     "stats": _stats, "implied": _ti}
                _new["log"] = (res.get("log") or []) + [
                    f"[precedent] {len(_keep)} of {len(_deals)} deals in the median, recomputed."]
                export_run(_new)
                st.session_state["res"] = _new
                snapshot_save(_new)
                st.session_state["precedent_ver"] = _pv + 1  # fresh editor next render
                st.rerun()
        _n_drop = len(_deals) - sum(1 for x in _applied if x)
        if _n_drop:
            st.caption(f"{_n_drop} deal(s) currently excluded from the median. "
                       "Tick them back in and recompute to include them again.")
        ti = res["precedent"]["implied"]
        if ti:
            st.markdown("##### Implied value from deal multiples")
            st.dataframe(pd.DataFrame([{"Multiple": t.get("label", t.get("multiple")),
                                        "Median": round(t["median"], 2) if t.get("median") else None,
                                        "Base EV": round((t.get("base_ev") or 0) / SCALE)}
                                       for t in ti]),
                         use_container_width=True, hide_index=True,
                         column_config={"Base EV":
                                        st.column_config.NumberColumn(f"Base EV (m {ccy})",
                                                                      format="localized")})
        st.caption("Deal multiples typically include a control premium, so read them as an "
                   "upper anchor versus trading comps.")
    else:
        st.info("No matching deals for this sector (or no target financials entered).")

# --- AI Match ---
with tabs[7]:
    if "match_score" in df.columns and df["match_score"].notna().any():
        ai_cols = [c for c in ["symbol", "name", "match_score", "match_rationale",
                               "business_profile"] if c in df.columns]
        st.dataframe(df[ai_cols].sort_values("match_score", ascending=False),
                     use_container_width=True, hide_index=True,
                     column_config={
                         "symbol": st.column_config.TextColumn("Ticker"),
                         "name": st.column_config.TextColumn("Company"),
                         "match_score": st.column_config.ProgressColumn(
                             "Fit", min_value=0, max_value=10, format="%d"),
                         "match_rationale": st.column_config.TextColumn("Rationale",
                                                                        width="large"),
                         "business_profile": st.column_config.TextColumn("Business profile",
                                                                         width="large")})
    else:
        st.info("Describe the target's business model in the sidebar to unlock AI matching.")

# --- Methodology ---
with tabs[8]:
    st.markdown("##### Trading comparables")
    st.markdown(
        "Peer financials are pulled per ticker and converted into the target currency. "
        "Multiples computed: EV/Sales, EV/EBITDA, EV/EBIT and P/E. EV-based multiples are "
        "capital-structure neutral, so they stay comparable across peers with different "
        "leverage; P/E does not, which is why it is reported but never leads.")
    st.latex(r"\text{EV} = \text{Market cap} + \text{Net debt}"
             r"\qquad \text{Implied EV} = \text{Target metric} \times \text{Peer multiple}")
    st.markdown(
        "The low/base/high band comes from Q1, median and Q3 of the peer set. Outliers are "
        "flagged with the IQR rule and stay visible in the long list, excluded only from "
        "the summary statistics. The median is used rather than the mean because multiple "
        "distributions are right-skewed and one 45x peer would drag a mean upwards.")
    st.latex(r"\text{outlier if } x < Q_1 - 1.5\,\text{IQR}"
             r"\quad\text{or}\quad x > Q_3 + 1.5\,\text{IQR},"
             r"\qquad \text{IQR} = Q_3 - Q_1")

    st.markdown("##### WACC and DCF")
    st.markdown(
        "Observed peer betas mix business risk with the financial risk of each peer's own "
        "debt. They are therefore unlevered to asset betas, the median is taken, and that "
        "median is relevered at the target's capital structure before CAPM.")
    st.latex(r"\beta_U = \frac{\beta_L}{1 + (1 - t)\,D/E}"
             r"\qquad \beta_{L,\text{target}} = \beta_U\left(1 + (1-t)\,(D/E)_{\text{target}}\right)")
    st.latex(r"k_e = r_f + \beta_{L,\text{target}}\,\text{MRP} + \text{size premium}"
             r"\qquad \text{WACC} = \frac{E}{V}k_e + \frac{D}{V}k_d(1-t)")
    st.markdown(
        "Cost of debt enters after tax because interest is deductible. A 5-year free cash "
        "flow projection plus a Gordon-growth terminal value is then discounted at that "
        "WACC. Terminal value usually carries 70% to 80% of enterprise value, which is why "
        "the sensitivity grid varies WACC against terminal growth: the result depends on "
        "the spread between them.")
    st.latex(r"\text{TV} = \frac{\text{FCF}_n (1+g)}{\text{WACC} - g}"
             r"\qquad \text{EV} = \sum_{i=1}^{n}\frac{\text{FCF}_i}{(1+\text{WACC})^i}"
             r" + \frac{\text{TV}}{(1+\text{WACC})^n}")
    st.markdown(
        "Peers with D/E above 200% distort the levered beta, so the run flags them and "
        "reports a cleaned WACC alongside the headline figure. Terminal growth above "
        "long-run GDP growth (roughly 2% to 2.5%) is an assumption error, not a bullish "
        "case: the company would eventually outgrow the economy.")

    st.markdown("##### Precedent transactions")
    st.markdown(
        "Deal multiples come from a local CSV (your own deal research) or a live FMP feed, "
        "filtered to the selected sector where the source tags one. They typically embed a "
        "control premium of roughly 20% to 40%, so they anchor the upper end of the "
        "football field and are never averaged together with trading comps.")

    st.markdown("##### AI peer match")
    st.markdown(
        "Gemini scores each peer's business-model similarity to the target description from "
        "1 to 10 with a written rationale. Scores of 6 or higher refine the beta pool for "
        "WACC, scores of 7 or higher pre-select the short list. The model ranks "
        "comparability only. It never reads, writes or adjusts a financial figure, so every "
        "number in the output is traceable to the data source.")

    st.markdown("##### Convergence check")
    st.markdown(
        "Where both exist, the DCF enterprise value is compared against the comps median. "
        "DCF reflects the assumptions entered here, comps reflect what the market pays "
        "today. A gap above 50% triggers a warning: neither number is wrong on its own, but "
        "one of the two needs explaining before the range goes into a deliverable.")

    st.markdown("##### Data and caching")
    st.markdown(
        f"Yahoo Finance fundamentals with a per-ticker cache, FX rates cached for 24 hours, "
        f"S&P 500 and STOXX 600 constituents for sector auto-discovery, optional FMP API for "
        f"deals. Every run writes `output/<target>_<timestamp>/` containing a formatted "
        f"Excel workbook and an optional PowerPoint deck. Version {APP_VERSION}.")

# ---------------------------------------------------------------- downloads
st.divider()
dl1, dl2 = st.columns(2)
xlsx_path = res.get("xlsx_path")
if xlsx_path and Path(xlsx_path).exists():
    dl1.download_button("Download Excel workbook", Path(xlsx_path).read_bytes(),
                        file_name=Path(xlsx_path).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
pptx_path = res.get("pptx_path")
if pptx_path and Path(pptx_path).exists():
    dl2.download_button("Download PowerPoint deck", Path(pptx_path).read_bytes(),
                        file_name=Path(pptx_path).name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True)
elif res.get("pptx_error"):
    dl2.caption(f"PowerPoint error: {res['pptx_error']}")
else:
    dl2.caption("PowerPoint needs target financials and the checkbox in Setup.")

st.markdown(
    f"<div class='pg-footer'>Peer Group Screening <b>v{APP_VERSION}</b> · "
    "Data: Yahoo Finance, FMP · Same backend as the desktop tool · "
    "For discussion purposes only. Not investment advice.</div>",
    unsafe_allow_html=True)
