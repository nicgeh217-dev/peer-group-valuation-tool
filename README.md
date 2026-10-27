# Peer Group Valuation Tool

An end-to-end company-valuation prototype: from a sector or ticker selection it builds a
full peer-group analysis — **Trading Comparables, a WACC-based DCF, Precedent
Transactions, and an AI-assisted peer match** — and exports a formula-driven Excel model
plus a PowerPoint deck.

Built to learn the comparable-companies and DCF workflow end-to-end. It is a **prototype
on free data**, not a production system — but every number is traceable, and the Excel
output is a live model, not a static export.

> **Honest note:** this was built AI-assisted. The point of the project was to understand
> the valuation methodology — unlever/relever, CAPM, control premium, IQR outliers,
> triangulation — well enough to implement it, not to hand-write every line.

---

## What it does

- **Trading comparables** — EV/Sales, EV/EBITDA, EV/EBIT, P/E; IQR outlier flagging
  (flagged, not deleted); quartile statistics; implied EV and equity range.
- **DCF on a peer-derived WACC** — peer betas unlevered (Hamada) and relevered to the
  target capital structure, cost of equity via CAPM; a 5-year FCF projection with an
  explicit **cash-flow bridge** (FCFF = EBIT·(1−t) + D&A − Capex − ΔNWC), Gordon-growth
  terminal value, and a sensitivity grid.
- **Precedent transactions** — deal multiples from a local CSV (or an optional live feed),
  curatable deal-by-deal; read as the control-premium anchor above trading comps.
- **AI peer match (optional)** — scores each peer's business-model fit to the target
  (1–10) with a written rationale. It ranks comparability only; it never touches the
  figures, and it can be switched off entirely.
- **Deliverables** — a fully **formula-driven** Excel workbook (change an input cell and
  the DCF recomputes) and a PowerPoint deck.
- **Analyst workflow** — universe → long list → filter/outlier → short list (per-row
  override) → recompute → football field with a DCF-vs-comps convergence check.

---

## Quickstart

```bash
git clone <this-repo>
cd peer-group-valuation-tool
pip install -r requirements.txt
streamlit run app.py
```

Then in the sidebar: **Load demo case → Run**. Or **Load last results** to see a saved
run instantly, with no network call.

### No API key needed to try it

- Screening, comps, WACC/DCF and precedent transactions all run on **free Yahoo Finance
  data — no key required**.
- The **AI peer match is optional**. To enable it, copy `.env.example` to `.env` and add a
  Gemini key (`GEMINI_API_KEY`). Without it, the AI toggle stays off and everything else
  works.
- The shipped demo snapshot lets **Load last results** show a complete valuation offline,
  straight after cloning.

---

## Tech stack

Python · Streamlit (web UI) · pandas · plotly · openpyxl · python-pptx · yfinance.
Data sources sit behind a swappable adapter interface, prepared for professional
providers (Capital IQ, Mergermarket) without rewriting the core.

## Tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

---

## Limitations (stated honestly)

- **Yahoo Finance** is free but has no SLA and is not licensed for commercial deliverables
  — fine for a prototype, a professional data source for production.
- Sector taxonomies differ (ICB in Europe vs GICS in the US), so cross-region peer sets
  are not one-to-one.
- Peer selection remains an analyst judgment; the AI match is an aid, not a replacement.

*For discussion purposes only. Not investment advice.*
