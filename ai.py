"""AI-powered business model matching via Google Gemini.

Sends target description + all peer business summaries in ONE batch call.
Gemini returns structured JSON with score (1-10) and rationale per peer.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

from paths import APP_DIR, CACHE_DIR
CACHE_DIR.mkdir(exist_ok=True)


def _load_api_key() -> str | None:
    """Read GEMINI_API_KEY from .env file or environment variable."""
    key = os.environ.get("GEMINI_API_KEY")
    if key:
        return key.strip()
    env_path = APP_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("GEMINI_API_KEY"):
                _, _, val = line.partition("=")
                return val.strip().strip('"').strip("'")
    return None


def is_available() -> bool:
    return _load_api_key() is not None


PROMPT_TEMPLATE = """Du bist M&A-Analyst und beurteilst, wie gut Peer-Unternehmen zum Target-Geschäftsmodell passen.

# Target-Beschreibung
{target_description}

# Peer-Unternehmen
{peer_block}

# Aufgabe
Für JEDES Peer-Unternehmen oben liefere:

1. **score** (1-10): wie gut Geschäftsmodell zum Target passt

   PRIORITÄTS-HIERARCHIE (in dieser Reihenfolge anwenden):
   a) STRUKTUR-MATCH (Wichtigster Faktor): Vertical SaaS ≠ Horizontal SaaS, auch bei gleicher Kundengröße.
      Wenn Target Vertical SaaS ist: andere Vertical SaaS = hoch, Horizontal Tools = niedrig — EGAL ob Kundensegment passt.
   b) BUSINESS-MODEL-MECHANIK: Erlösmodell (Subscription/Usage/Lizenz), Sticky-Level, Implementation-Tiefe.
   c) Branche/Industrie: nachrangig zur Struktur. Vertical SaaS in anderer Industrie schlägt Horizontal SaaS in gleicher Industrie.
   d) Kundensegment-Größe: letzte Stelle — nicht aufwerten nur weil "auch KMU".

   - 10 = Direkter Wettbewerber, exakt gleiche Vertical, gleiches Produkt
   - 7-9 = Gleiche Struktur (z.B. Vertical SaaS) in anderer Industrie ABER mit vergleichbarem Playbook
   - 4-6 = Verwandte Struktur, aber wichtiger Unterschied (z.B. Horizontal statt Vertical bei sonst-gleichem Kunden)
   - 1-3 = Andere Struktur ODER völlig andere Mechanik (z.B. Infrastruktur, Dev-Tools, anderes Erlösmodell)

2. **rationale**: max. 15 Wörter, warum dieser Score

3. **profile**: Workshop-Snapshot mit 5 kurzen Bullet-Werten:
   - core: Was macht Firma (3-6 Wörter, z.B. "B2B SaaS Projektmanagement")
   - customer: Kundensegment (z.B. "Mittelstand + Enterprise", "KMU", "Großkonzerne")
   - gtm: Vertriebsmodell (z.B. "Direct Sales", "Self-Serve + Inside Sales", "Channel/Partner")
   - geo: Geografie-Schwerpunkt (z.B. "USA + Europa", "Global", "DACH")
   - revenue: Erlösmodell (z.B. "Subscription (ARR)", "Lizenz + Wartung", "Usage-based", "Services")

# Output-Format
NUR valides JSON, keine Erklärung davor/danach:

{{
  "scores": [
    {{
      "ticker": "XYZ.DE",
      "score": 8,
      "rationale": "Direkter SaaS-Wettbewerber im Buchhaltungssegment.",
      "profile": {{
        "core": "B2B SaaS Accounting",
        "customer": "Mittelstand DACH",
        "gtm": "Direct Sales + Channel",
        "geo": "Europa",
        "revenue": "Subscription (ARR)"
      }}
    }},
    ...
  ]
}}"""


def score_peers(
    target_description: str,
    peers: list[dict],
    model: str = "gemini-2.5-flash",
    batch_size: int = 20,
) -> dict[str, dict]:
    """
    target_description: free text from the analyst
    peers: list of dicts with keys: symbol, name, sector, industry, business_summary

    Returns: {ticker: {"score": int, "rationale": str, "business_profile": str}}
    Large universes are split into batches of `batch_size` to keep response time per call low.
    """
    key = _load_api_key()
    if not key or not target_description.strip() or not peers:
        return {}

    # Filter peers that actually have a business summary
    peers_with_summary = [p for p in peers if (p.get("business_summary") or "").strip()]
    if not peers_with_summary:
        return {}

    # Batch large universes
    if len(peers_with_summary) > batch_size:
        merged = {}
        for i in range(0, len(peers_with_summary), batch_size):
            chunk = peers_with_summary[i:i + batch_size]
            print(f"      AI batch {i//batch_size + 1}: {len(chunk)} peers")
            merged.update(_score_batch(target_description, chunk, key, model))
        return merged
    return _score_batch(target_description, peers_with_summary, key, model)


def _score_batch(target_description: str, peers_with_summary: list[dict], key: str, model: str) -> dict[str, dict]:

    # Cache key based on target description + tickers + summaries
    cache_input = (
        target_description.strip()
        + "|"
        + "|".join(f"{p['symbol']}:{(p.get('business_summary') or '')[:200]}" for p in peers_with_summary)
    )
    cache_key = hashlib.md5(cache_input.encode("utf-8")).hexdigest()
    cache_file = CACHE_DIR / f"ai_scores_{cache_key}.json"
    if cache_file.exists():
        try:
            return json.loads(cache_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass

    # Build peer block
    peer_lines = []
    for p in peers_with_summary:
        summary = (p.get("business_summary") or "")[:800]  # trim to control prompt size
        line = (
            f"Ticker: {p['symbol']}\n"
            f"Name: {p.get('name') or p['symbol']}\n"
            f"Sector / Industry: {p.get('sector') or '-'} / {p.get('industry') or '-'}\n"
            f"Business Summary: {summary}\n"
        )
        peer_lines.append(line)
    peer_block = "\n---\n".join(peer_lines)

    prompt = PROMPT_TEMPLATE.format(
        target_description=target_description.strip(),
        peer_block=peer_block,
    )

    # Direct REST call — avoids google.genai SDK namespace-package issues on Python 3.14
    import requests, time
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "responseMimeType": "application/json",
        },
    }

    # Retry on timeout AND on transient server errors (503 overload, 429 rate limit, 500)
    RETRY_CODES = {429, 500, 502, 503, 504}
    resp = None
    last_err = None
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            resp = requests.post(
                url,
                params={"key": key},
                json=payload,
                timeout=180,
                headers={"Content-Type": "application/json"},
            )
        except requests.exceptions.ReadTimeout as e:
            last_err = e
            print(f"      Gemini timeout (Versuch {attempt + 1}/{max_attempts}), retry…")
            time.sleep(2 * (attempt + 1))
            continue
        if resp.status_code in RETRY_CODES and attempt < max_attempts - 1:
            wait = 3 * (attempt + 1)  # backoff: 3,6,9,12s
            print(f"      Gemini {resp.status_code} (überlastet) — warte {wait}s, retry {attempt + 1}/{max_attempts}…")
            time.sleep(wait)
            continue
        break
    if resp is None:
        raise RuntimeError(f"Gemini nicht erreichbar nach {max_attempts} Versuchen: {last_err}")
    if resp.status_code != 200:
        raise RuntimeError(f"Gemini API {resp.status_code}: {resp.text[:200]}")

    body = resp.json()
    try:
        text = body["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError, TypeError):
        return {}

    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return {}

    out = {}
    for row in data.get("scores", []):
        ticker = row.get("ticker")
        if not ticker:
            continue
        profile = row.get("profile") or {}
        # Compact multi-line format for Excel
        profile_lines = []
        for key, label in [
            ("core",     "Was"),
            ("customer", "Kunden"),
            ("gtm",      "GTM"),
            ("geo",      "Geo"),
            ("revenue",  "Modell"),
        ]:
            v = (profile.get(key) or "").strip()
            if v:
                profile_lines.append(f"• {label}: {v}")
        out[ticker] = {
            "score":             row.get("score"),
            "rationale":         (row.get("rationale") or "").strip(),
            "business_profile":  "\n".join(profile_lines),
        }

    try:
        cache_file.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

    return out
