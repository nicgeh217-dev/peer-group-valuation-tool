"""Excel export: Long List + Summary tabs."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTLIER_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
EXCLUDED_FILL = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")

def _long_list_cols(target_ccy: str):
    """Column spec. Money columns include the target currency in the header."""
    money = f" ({target_ccy})"
    return [
        ("symbol", "Ticker", None),
        ("name", "Company", None),
        ("match_score", "Match", "0"),
        ("match_rationale", "Match Rationale", None),
        ("business_profile", "Business Profile", None),
        ("country", "Country", None),
        ("sector", "Sector", None),
        ("industry", "Industry", None),
        ("currency", "Reporting Ccy", None),
        ("fx_rate", "FX Rate", "0.0000"),
        ("revenue", f"Revenue{money}", "#,##0"),
        ("ebitda", f"EBITDA{money}", "#,##0"),
        ("ebit", f"EBIT{money}", "#,##0"),
        ("net_income", f"Net Income{money}", "#,##0"),
        ("ebitda_margin", "EBITDA Margin", "0.0%"),
        ("ebit_margin", "EBIT Margin", "0.0%"),
        ("market_cap", f"Market Cap{money}", "#,##0"),
        ("net_debt", f"Net Debt{money}", "#,##0"),
        ("enterprise_value", f"EV{money}", "#,##0"),
        ("ev_sales", "EV/Sales", "0.00x"),
        ("ev_ebitda", "EV/EBITDA", "0.00x"),
        ("ev_ebit", "EV/EBIT", "0.00x"),
        ("pe", "P/E", "0.00x"),
        ("fiscal_date", "Fiscal Date", None),
        ("excluded_reason", "Excluded", None),
        ("outlier_flags", "Outlier Flags", None),
    ]


def _autosize(ws):
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 10
        for cell in ws[letter]:
            v = cell.value
            if v is None:
                continue
            max_len = max(max_len, min(len(str(v)) + 2, 40))
        ws.column_dimensions[letter].width = max_len


def export(
    long_list_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    target_meta: dict,
    output_path: Path,
    skipped: list[tuple[str, str]] | None = None,
    implied: list[dict] | None = None,
    target_inputs: dict | None = None,
    wacc_result: dict | None = None,
    dcf_result: dict | None = None,
    dcf_sens: list[dict] | None = None,
    convergence: dict | None = None,
    precedent: dict | None = None,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Determine target currency from meta. The web app passes "Target currency",
    # the desktop app (gui.py) passes "Zielwährung" — accept both so the Excel labels
    # never silently fall back to USD on a non-USD run.
    target_ccy = target_meta.get("Target currency") or target_meta.get("Zielwährung") or "USD"
    cols_spec = _long_list_cols(target_ccy)

    # ---- Long List tab ----
    ws = wb.active
    ws.title = "Long List"

    headers = [label for _, label, _ in cols_spec]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for _, row in long_list_df.iterrows():
        values = []
        for key, _, _ in cols_spec:
            v = row.get(key)
            if pd.isna(v) if not isinstance(v, (str, list)) else False:
                v = None
            values.append(v)
        ws.append(values)

    # Match-score color scale (light red→white→green) per score
    def _match_fill(score):
        if score is None or pd.isna(score):
            return None
        try:
            s = int(score)
        except (ValueError, TypeError):
            return None
        # 10→deep green, 5→light, 1→red
        if s >= 8:
            return PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        if s >= 6:
            return PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
        if s >= 4:
            return PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        return PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")

    # Number formats + row highlighting
    for r_idx, (_, row) in enumerate(long_list_df.iterrows(), start=2):
        excluded = bool(row.get("excluded_reason"))
        outlier = bool(row.get("outlier_flags"))
        match_score = row.get("match_score")
        for c_idx, (key, _, fmt) in enumerate(cols_spec, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if fmt:
                cell.number_format = fmt
            # Row-level fill first
            if excluded:
                cell.fill = EXCLUDED_FILL
            elif outlier:
                cell.fill = OUTLIER_FILL
            # Then override match-score cell with its specific color
            if key == "match_score":
                mf = _match_fill(match_score)
                if mf:
                    cell.fill = mf
            # Wrap multi-line profile cell
            if key == "business_profile":
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "C2"
    _autosize(ws)

    # ---- Summary tab ----
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Peer Group Summary (outliers excluded)"
    ws2["A1"].font = Font(bold=True, size=14)

    ws2["A3"] = "Target"
    ws2["A3"].font = Font(bold=True)
    row = 4
    for k, v in target_meta.items():
        ws2.cell(row=row, column=1, value=str(k))
        ws2.cell(row=row, column=2, value=str(v) if v is not None else "")
        row += 1
    row += 1

    ws2.cell(row=row, column=1, value="Multiple Statistics").font = Font(bold=True)
    row += 1
    stat_headers = ["Metric", "N", "Mean", "Median", "Q1", "Q3", "Min", "Max"]
    for i, h in enumerate(stat_headers, start=1):
        c = ws2.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    row += 1

    for _, s in summary_df.iterrows():
        ws2.cell(row=row, column=1, value=s["metric"])
        ws2.cell(row=row, column=2, value=s["n"])
        for i, k in enumerate(["mean", "median", "q1", "q3", "min", "max"], start=3):
            val = s[k]
            cell = ws2.cell(row=row, column=i, value=val)
            if s["metric"].endswith("_margin"):
                cell.number_format = "0.0%"
            else:
                cell.number_format = "0.00x"
        row += 1

    _autosize(ws2)

    # ---- Business Models tab (if AI scored anything) ----
    if "match_score" in long_list_df.columns and long_list_df["match_score"].notna().any():
        _write_business_models_tab(wb, long_list_df)

    # ---- Valuation tab (only if target inputs were provided) ----
    if implied and target_inputs:
        _write_valuation_tab(wb, implied, target_inputs, target_ccy)

    # ---- WACC + DCF tabs ----
    if wacc_result:
        _write_wacc_tab(wb, wacc_result, target_ccy)
    if dcf_result and dcf_result.get("enterprise_value"):
        _write_dcf_tab(wb, dcf_result, dcf_sens, wacc_result, target_ccy, convergence)

    # ---- Precedent Transactions tab ----
    if precedent and precedent.get("implied"):
        _write_precedent_tab(wb, precedent, target_inputs, target_ccy)

    # ---- Football Field (after valuation + dcf so it can include DCF + precedent bars) ----
    if implied and target_inputs:
        # Peer count for quartile-reliability warning
        n_peers = 0
        try:
            ns = pd.to_numeric(summary_df["n"], errors="coerce").dropna()
            n_peers = int(ns.max()) if len(ns) else 0
        except Exception:
            n_peers = 0
        _write_football_field(wb, implied, target_inputs, target_ccy, dcf_result, n_peers, precedent)

    # ---- Skipped tab (only if anything was skipped) ----
    if skipped:
        ws3 = wb.create_sheet("Skipped")
        ws3["A1"] = "Übersprungene Ticker (keine Equity-Daten)"
        ws3["A1"].font = Font(bold=True, size=12)
        ws3["A3"] = "Ticker"; ws3["B3"] = "Grund"
        for c in (ws3["A3"], ws3["B3"]):
            c.fill = HEADER_FILL; c.font = HEADER_FONT
        for i, (sym, reason) in enumerate(skipped, start=4):
            ws3.cell(row=i, column=1, value=sym)
            ws3.cell(row=i, column=2, value=reason)
        _autosize(ws3)

    wb.save(output_path)
    return output_path


def _write_business_models_tab(wb, df: pd.DataFrame):
    """Lists each peer with full business summary, AI score, rationale. Sorted by score desc."""
    ws = wb.create_sheet("Business Models")
    ws["A1"] = "Business Models — AI Match Score"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Sortiert nach Match-Score absteigend. Color-Code: dunkelgrün ≥8, hellgrün ≥6, gelb ≥4, rot <4."
    ws["A2"].font = Font(italic=True, color="666666")

    headers = ["Ticker", "Company", "Score", "Rationale", "Business Profile", "Industry", "Business Summary"]
    row = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")

    def _fill_for(s):
        if s is None or pd.isna(s):
            return None
        try:
            s = int(s)
        except (ValueError, TypeError):
            return None
        if s >= 8: return PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        if s >= 6: return PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
        if s >= 4: return PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        return PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")

    sorted_df = df.copy()
    sorted_df["_s"] = pd.to_numeric(sorted_df.get("match_score"), errors="coerce").fillna(-1)
    sorted_df = sorted_df.sort_values("_s", ascending=False)

    row = 5
    for _, r in sorted_df.iterrows():
        ws.cell(row=row, column=1, value=r.get("symbol"))
        ws.cell(row=row, column=2, value=r.get("name"))
        score = r.get("match_score")
        c = ws.cell(row=row, column=3, value=score if pd.notna(score) else None)
        c.alignment = Alignment(horizontal="center")
        fill = _fill_for(score)
        if fill:
            c.fill = fill
        ws.cell(row=row, column=4, value=r.get("match_rationale"))
        ws.cell(row=row, column=4).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row, column=5, value=r.get("business_profile"))
        ws.cell(row=row, column=5).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row, column=6, value=r.get("industry"))
        ws.cell(row=row, column=7, value=r.get("business_summary"))
        ws.cell(row=row, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    # Column widths
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 45
    ws.column_dimensions["E"].width = 35
    ws.column_dimensions["F"].width = 25
    ws.column_dimensions["G"].width = 70


def _write_football_field(wb, implied: list[dict], target_inputs: dict, target_ccy: str,
                          dcf_result: dict | None = None, n_peers: int = 0,
                          precedent: dict | None = None):
    """Football Field chart: floating bar per method showing value range.
    Trading Comps multiples + (if available) a DCF bar.

    Technique: stacked horizontal bar.
      Series 1 = low (Q1)        → no fill (invisible pedestal)
      Series 2 = median - low    → light bar (Q1 → Median)
      Series 3 = high - median   → darker bar (Median → Q3)
    The boundary between the two visible segments marks the Median.
    """
    from openpyxl.chart.label import DataLabelList

    ws = wb.create_sheet("Football Field")
    money = f"({target_ccy})"
    scale = 1_000_000  # display in millions
    unit = f"Mio {target_ccy}"

    ws["A1"] = "Football Field — Implied Enterprise Value"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Jeder Balken = eine Methode. Spanne Q1→Q3, Trennlinie = Median. Werte in {unit}."
    ws["A2"].font = Font(italic=True, color="666666")

    # --- Reliability warning if too few peers ---
    warn_row = None
    if n_peers and n_peers < 4:
        ws["A3"] = (f"⚠ Nur {n_peers} Peer(s) in der Statistik — Quartile (Q1/Q3) wenig aussagekräftig. "
                    f"Für belastbare Spannen ≥4 Peers wählen.")
        ws["A3"].font = Font(color="C0392B", bold=True, size=10)
        warn_row = 3

    # --- Data block ---
    # Cols: Label | Q1(pedestal) | Range(High-Q1) | Median | Low | High
    hdr_row = 5
    headers = ["Methode", "Sockel (Q1)", "Range", "Median", "Low (Q1)", "High (Q3)"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=hdr_row, column=i, value=h)
        c.font = HEADER_FONT; c.fill = HEADER_FILL

    data_start = hdr_row + 1
    r = data_start
    all_highs = []

    def _bar_row(label, low, base, high):
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=round(low, 1))            # pedestal (invisible)
        ws.cell(row=r, column=3, value=round(high - low, 1))     # visible range
        ws.cell(row=r, column=4, value=round(base, 1))           # median (reference)
        ws.cell(row=r, column=5, value=round(low, 1))
        ws.cell(row=r, column=6, value=round(high, 1))
        for col in range(2, 7):
            ws.cell(row=r, column=col).number_format = "#,##0"
        all_highs.append(high)

    for item in implied:
        low = (item.get("low_ev") or 0) / scale
        high = (item.get("high_ev") or 0) / scale
        base = (item.get("base_ev") or 0) / scale
        _bar_row(item["label"], low, base, high)
        r += 1

    if dcf_result and dcf_result.get("enterprise_value"):
        dcf_ev = dcf_result["enterprise_value"] / scale
        _bar_row("DCF (Base ±10%)", dcf_ev * 0.9, dcf_ev, dcf_ev * 1.1)
        r += 1

    # Precedent transactions bar — median EV/EBITDA across deals (control premium embedded)
    if precedent and precedent.get("implied"):
        tx = precedent["implied"]
        # Use EV/EBITDA deal if available, else first available
        pick = next((t for t in tx if t["multiple"] == "ev_ebitda"), tx[0])
        if pick.get("base_ev"):
            low = (pick.get("low_ev") or pick["base_ev"]) / scale
            high = (pick.get("high_ev") or pick["base_ev"]) / scale
            _bar_row("Precedent Tx", low, pick["base_ev"] / scale, high)
            r += 1

    data_end = r - 1

    # --- Build floating horizontal bar chart (2 series: invisible pedestal + visible range) ---
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"Implied EV Range ({unit})"
    chart.height = 9
    chart.width = 24

    data_ref = Reference(ws, min_col=2, max_col=3, min_row=hdr_row, max_row=data_end)
    cats = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)

    pedestal, rangebar = chart.series[0], chart.series[1]
    pedestal.graphicalProperties.noFill = True
    pedestal.graphicalProperties.line.noFill = True
    rangebar.graphicalProperties.solidFill = "4472C4"
    rangebar.graphicalProperties.line.solidFill = "2F5597"

    chart.legend = None
    ws.add_chart(chart, "H5")

    # --- Labelled value table beside chart: Methode | Low (Q1) | Median | High (Q3) ---
    ann_row = data_end + 3
    ws.cell(row=ann_row, column=1, value=f"Bewertungsspanne je Methode ({unit})").font = Font(bold=True, size=11)
    ann_row += 1
    tbl_hdr = ["Methode", "Low (Q1)", "Median", "High (Q3)"]
    for i, h in enumerate(tbl_hdr, start=1):
        c = ws.cell(row=ann_row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    ann_row += 1

    def _tbl_row(label, low, base, high, highlight=False):
        ws.cell(row=ann_row, column=1, value=label)
        for col, val in [(2, low), (3, base), (4, high)]:
            c = ws.cell(row=ann_row, column=col, value=round(val / scale, 0))
            c.number_format = "#,##0"
        # Median cell emphasized
        ws.cell(row=ann_row, column=3).font = Font(bold=True)
        if highlight:
            for col in range(1, 5):
                ws.cell(row=ann_row, column=col).fill = PatternFill(
                    start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

    for item in implied:
        _tbl_row(item["label"], item.get("low_ev") or 0, item.get("base_ev") or 0, item.get("high_ev") or 0)
        ann_row += 1
    if dcf_result and dcf_result.get("enterprise_value"):
        ev = dcf_result["enterprise_value"]
        _tbl_row("DCF", ev * 0.9, ev, ev * 1.1, highlight=True)
        ann_row += 1
    if precedent and precedent.get("implied"):
        pick = next((t for t in precedent["implied"] if t["multiple"] == "ev_ebitda"), precedent["implied"][0])
        if pick.get("base_ev"):
            _tbl_row("Precedent Tx", pick.get("low_ev") or pick["base_ev"], pick["base_ev"],
                     pick.get("high_ev") or pick["base_ev"], highlight=True)
            ann_row += 1

    # Overall range
    ann_row += 1
    lows = [(i.get("low_ev") or 0) for i in implied]
    highs = [(i.get("high_ev") or 0) for i in implied]
    if dcf_result and dcf_result.get("enterprise_value"):
        lows.append(dcf_result["enterprise_value"] * 0.9)
        highs.append(dcf_result["enterprise_value"] * 1.1)
    if lows and highs:
        ws.cell(row=ann_row, column=1, value="Gesamt-Spanne").font = Font(bold=True)
        c = ws.cell(row=ann_row, column=2, value=round(min(lows) / scale, 0)); c.number_format = "#,##0"
        ws.cell(row=ann_row, column=3, value="bis")
        c = ws.cell(row=ann_row, column=4, value=round(max(highs) / scale, 0)); c.number_format = "#,##0"

    ws.column_dimensions["A"].width = 18
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 11


def _write_valuation_tab(wb, implied: list[dict], target_inputs: dict, target_ccy: str):
    """Builds the Implied Valuation tab: target inputs, EV table, Equity table, summary range.

    Formula-driven: the target financials are input cells, and every implied value is
    a live Excel formula (Implied EV = operand × peer multiple; Equity = EV − Net Debt).
    Change an input or a multiple and the whole tab recomputes — this is a model, not a
    static report. The peer multiples (Q1/Median/Q3) are the tab's own inputs and stay
    as values.
    """
    from valuation import summary_range

    ws = wb.create_sheet("Valuation")
    money = f"({target_ccy})"

    ws["A1"] = "Implied Valuation"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Trading-Comps-Bewertung: Implied EV = Target-Operand × Peer-Multiple. Blaue Zellen = Eingaben, Rest = Formeln."
    ws["A2"].font = Font(italic=True, color="666666")

    INPUT_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    BASE_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

    # --- Target inputs section (these are the input cells everything references) ---
    row = 4
    ws.cell(row=row, column=1, value="Target Inputs (Eingaben)").font = Font(bold=True, size=11)
    row += 1
    label_map = [
        ("revenue",    "Revenue"),
        ("ebitda",     "EBITDA"),
        ("ebit",       "EBIT"),
        ("net_income", "Net Income"),
        ("net_debt",   "Net Debt"),
    ]
    input_rows: dict[str, int] = {}
    for key, lbl in label_map:
        val = target_inputs.get(key)
        # Net debt feeds the equity bridge as a formula operand, so a missing value
        # must be a numeric 0, not an em dash, or the bridge formula would error.
        if val is None and key == "net_debt":
            val = 0
        ws.cell(row=row, column=1, value=f"{lbl} {money}")
        c = ws.cell(row=row, column=2, value=val if val is not None else "—")
        c.number_format = "#,##0"
        if val is not None:
            c.fill = INPUT_FILL
        input_rows[key] = row
        row += 1
    row += 1

    nd_ref = f"$B${input_rows['net_debt']}"
    operand_key = {"ev_sales": "revenue", "ev_ebitda": "ebitda",
                   "ev_ebit": "ebit", "pe": "net_income"}

    headers = ["Multiple", "Operand", "Peer Q1", "Peer Median", "Peer Q3",
               "Low (Q1)", "Base (Median)", "High (Q3)"]

    # --- Implied EV table ---
    ws.cell(row=row, column=1, value=f"Implied Enterprise Value {money}").font = Font(bold=True, size=11)
    row += 1
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    row += 1

    ev_rows: dict[str, int] = {}   # multiple -> EV table row, for the equity bridge
    for r in implied:
        op_row = input_rows.get(operand_key.get(r["multiple"], ""))
        op_ref = f"$B${op_row}" if op_row else None
        ws.cell(row=row, column=1, value=r["label"])
        # Operand as a live reference to the target input, so it moves with the input
        if op_ref:
            ws.cell(row=row, column=2, value=f"={op_ref}").number_format = "#,##0"
        else:
            ws.cell(row=row, column=2, value=r.get("operand_val")).number_format = "#,##0"
        for i, key in enumerate(["q1", "median", "q3"], start=3):
            c = ws.cell(row=row, column=i, value=r[key])
            c.number_format = "0.00x"
        # Implied EV formulas. EV-multiples: EV = operand × multiple.
        # P/E: operand × multiple is EQUITY, so bridge up to EV by adding net debt.
        mult_cols = {"low": "C", "base": "D", "high": "E"}
        for col_idx, band in ((6, "low"), (7, "base"), (8, "high")):
            mc = mult_cols[band]
            if op_ref and r["multiple"] == "pe":
                f = f"=$B${row}*{mc}{row}+{nd_ref}"
            elif op_ref:
                f = f"=$B${row}*{mc}{row}"
            else:
                f = None
            cell = ws.cell(row=row, column=col_idx)
            cell.value = f if f else r.get({"low": "low_ev", "base": "base_ev",
                                            "high": "high_ev"}[band])
            cell.number_format = "#,##0"
        ws.cell(row=row, column=7).fill = BASE_FILL
        ev_rows[r["multiple"]] = row
        row += 1

    row += 1
    # --- Implied Equity table ---
    ws.cell(row=row, column=1,
            value=f"Implied Equity Value {money}  (EV − Net Debt)").font = Font(bold=True, size=11)
    row += 1
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    row += 1

    for r in implied:
        op_row = input_rows.get(operand_key.get(r["multiple"], ""))
        op_ref = f"$B${op_row}" if op_row else None
        ev_row = ev_rows.get(r["multiple"])
        ws.cell(row=row, column=1, value=r["label"])
        if op_ref:
            ws.cell(row=row, column=2, value=f"={op_ref}").number_format = "#,##0"
        else:
            ws.cell(row=row, column=2, value=r.get("operand_val")).number_format = "#,##0"
        for i, key in enumerate(["q1", "median", "q3"], start=3):
            c = ws.cell(row=row, column=i, value=r[key])
            c.number_format = "0.00x"
        mult_cols = {"low": ("C", 6), "base": ("D", 7), "high": ("E", 8)}
        for band, (mc, col_idx) in mult_cols.items():
            if op_ref and r["multiple"] == "pe":
                # P/E gives equity directly: operand × multiple
                f = f"=$B${row}*{mc}{row}"
            elif op_ref and ev_row:
                # EV-multiples: equity = the matching EV cell − net debt
                ev_col = {"low": "F", "base": "G", "high": "H"}[band]
                f = f"={ev_col}{ev_row}-{nd_ref}"
            else:
                f = None
            cell = ws.cell(row=row, column=col_idx)
            cell.value = f if f else r.get({"low": "low_equity", "base": "base_equity",
                                            "high": "high_equity"}[band])
            cell.number_format = "#,##0"
        ws.cell(row=row, column=7).fill = BASE_FILL
        row += 1

    row += 1
    # --- Summary range (min/max across the formula cells above) ---
    rng = summary_range(implied)
    if rng and implied:
        ev_first, ev_last = min(ev_rows.values()), max(ev_rows.values())
        ws.cell(row=row, column=1, value="Valuation Range (über alle Multiples)").font = Font(bold=True, size=11)
        row += 1
        ws.cell(row=row, column=1, value=f"Enterprise Value {money}")
        ws.cell(row=row, column=2, value="Min:")
        c = ws.cell(row=row, column=3, value=f"=MIN(F{ev_first}:H{ev_last})"); c.number_format = "#,##0"
        ws.cell(row=row, column=4, value="Max:")
        c = ws.cell(row=row, column=5, value=f"=MAX(F{ev_first}:H{ev_last})"); c.number_format = "#,##0"
        row += 1
        # Equity range: the equity table sits directly below the EV table; its rows are
        # offset by a fixed gap, but MIN/MAX over the values is simplest kept static here.
        ws.cell(row=row, column=1, value=f"Equity Value {money}")
        ws.cell(row=row, column=2, value="Min:")
        c = ws.cell(row=row, column=3, value=rng.get("equity_min")); c.number_format = "#,##0"
        ws.cell(row=row, column=4, value="Max:")
        c = ws.cell(row=row, column=5, value=rng.get("equity_max")); c.number_format = "#,##0"
        row += 1

    _autosize(ws)


def _write_wacc_tab(wb, w: dict, target_ccy: str):
    """WACC breakdown: peer betas, unlever/relever, CAPM, weighted cost.

    Formula-driven. The rates and the target D/E are input cells; the unlever step,
    the median, the relever step and the whole CAPM chain are live Excel formulas that
    reference them. Change the tax rate or the target D/E and the WACC recomputes.
    """
    ws = wb.create_sheet("WACC")
    ws["A1"] = "WACC — Cost of Capital"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Peer Betas → unlevern → Median → relevern → CAPM → WACC. Blaue Zellen = Eingaben, Rest = Formeln."
    ws["A2"].font = Font(italic=True, color="666666")

    INPUT_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    # --- Assumptions block first, at fixed cells, so the peer table below can
    #     reference the tax cell in its unlever formula ---
    ar = 4
    ws.cell(row=ar, column=1, value="Annahmen (Eingaben)").font = Font(bold=True, size=11)
    ar += 1
    assum = [
        ("tax",  "Tax Rate", "0.00%"),
        ("rf",   "Risk-free Rate (Rf)", "0.00%"),
        ("mrp",  "Market Risk Premium (MRP)", "0.00%"),
        ("size_premium", "Size Premium", "0.00%"),
        ("kd",   "Cost of Debt (Kd)", "0.00%"),
        ("target_de", "Target D/E", "0.00"),
    ]
    arow: dict[str, int] = {}
    for key, lbl, fmt in assum:
        ws.cell(row=ar, column=1, value=lbl)
        c = ws.cell(row=ar, column=2, value=w.get(key, 0))
        c.number_format = fmt; c.fill = INPUT_FILL
        arow[key] = ar
        ar += 1
    TAX = f"$B${arow['tax']}"; RF = f"$B${arow['rf']}"; MRP = f"$B${arow['mrp']}"
    SIZE = f"$B${arow['size_premium']}"; KD = f"$B${arow['kd']}"; TDE = f"$B${arow['target_de']}"

    # --- Peer beta table ---
    row = ar + 1
    ws.cell(row=row, column=1, value="Peer Betas").font = Font(bold=True, size=11)
    row += 1
    headers = ["Ticker", "Company", "Beta (levered)", "D/E", "Beta (unlevered)"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
    row += 1
    warn_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    beta_u_first = row
    for p in w.get("peer_betas", []):
        ws.cell(row=row, column=1, value=p["symbol"])
        ws.cell(row=row, column=2, value=p.get("name"))
        ws.cell(row=row, column=3, value=round(p["beta_levered"], 3))
        c = ws.cell(row=row, column=4, value=p["de_ratio"]); c.number_format = "0.0%"
        # Unlever as a live formula: βu = βl / (1 + (1−Tax)·D/E)
        c = ws.cell(row=row, column=5, value=f"=C{row}/(1+(1-{TAX})*D{row})")
        c.number_format = "0.000"
        if p.get("de_outlier"):
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = warn_fill
            ws.cell(row=row, column=6, value="⚠ D/E >200% — Beta verzerrt")
        row += 1
    beta_u_last = row - 1

    # Diagnostic note if outliers present (kept as values — a filtered median is a
    # diagnostic, not part of the live chain)
    if w.get("n_de_outliers"):
        ws.cell(row=row, column=1,
                value=f"⚠ {w['n_de_outliers']} Peer(s) mit D/E >200% (rot). Distressed-Kapitalstruktur verzerrt Beta nach unten.").font = Font(color="C0392B", size=9, italic=True)
        row += 1
        if w.get("median_unlevered_beta_clean") is not None:
            ws.cell(row=row, column=1, value="Median Unlevered Beta (ohne D/E-Outlier)")
            ws.cell(row=row, column=2, value=round(w["median_unlevered_beta_clean"], 3))
            if w.get("wacc_clean"):
                ws.cell(row=row, column=3, value="→ WACC bereinigt:")
                c = ws.cell(row=row, column=4, value=w["wacc_clean"]); c.number_format = "0.00%"
            row += 1
    row += 1

    # --- Derived chain (all formulas referencing the inputs + the peer table) ---
    ws.cell(row=row, column=1, value="Berechnung").font = Font(bold=True, size=11)
    row += 1

    def _formula_line(lbl, formula, fmt):
        nonlocal row
        ws.cell(row=row, column=1, value=lbl)
        c = ws.cell(row=row, column=2, value=formula)
        c.number_format = fmt
        r = row
        row += 1
        return r

    r_medbu = _formula_line("Median Unlevered Beta",
                            f"=MEDIAN(E{beta_u_first}:E{beta_u_last})", "0.000")
    r_relev = _formula_line("Relevered Beta (Target)  βu·(1+(1−T)·D/E)",
                            f"=B{r_medbu}*(1+(1-{TAX})*{TDE})", "0.000")
    row += 1
    r_ke = _formula_line("Cost of Equity (Ke = Rf + β·MRP + Size)",
                         f"={RF}+B{r_relev}*{MRP}+{SIZE}", "0.00%")
    row += 1
    r_atkd = _formula_line("After-tax Kd  (Kd·(1−T))", f"={KD}*(1-{TAX})", "0.00%")
    r_we = _formula_line("Weight Equity (E/V = 1/(1+D/E))", f"=1/(1+{TDE})", "0.0%")
    r_wd = _formula_line("Weight Debt (D/V = D/E/(1+D/E))", f"={TDE}/(1+{TDE})", "0.0%")
    row += 1
    # WACC headline
    ws.cell(row=row, column=1, value="WACC = E/V·Ke + D/V·Kd·(1−T)").font = Font(bold=True, size=13)
    c = ws.cell(row=row, column=2, value=f"=B{r_we}*B{r_ke}+B{r_wd}*B{r_atkd}")
    c.number_format = "0.00%"; c.font = Font(bold=True, size=13)
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["E"].width = 18


def _write_dcf_tab(wb, d: dict, sens: list, w: dict, target_ccy: str, convergence: dict | None = None):
    """DCF: projected FCF, discounting, terminal value, EV + sensitivity grid."""
    ws = wb.create_sheet("DCF")
    money = f"({target_ccy})"
    ws["A1"] = "DCF — Discounted Cash Flow"
    ws["A1"].font = Font(bold=True, size=14)
    wacc = w.get("wacc") if w else None
    ws["A2"] = f"FCF projiziert, mit WACC {wacc:.2%} diskontiert, Terminal Value via Gordon Growth." if wacc else "DCF"
    ws["A2"].font = Font(italic=True, color="666666")

    # --- Convergence check banner ---
    row = 3
    if convergence and convergence.get("gap") is not None:
        gap = convergence["gap"]
        scale = 1e6
        msg = (f"Konvergenz-Check: DCF {convergence['dcf_ev']/scale:,.0f} vs "
               f"Comps-Median {convergence['comps_median_ev']/scale:,.0f} {money} — "
               f"Abweichung {gap:.0%}")
        c = ws.cell(row=row, column=1, value=msg)
        if gap > 0.5:
            c.value = "⚠ " + msg + "  →  Methoden divergieren stark, WACC/Annahmen prüfen."
            c.font = Font(color="C0392B", bold=True, size=10)
        else:
            c.value = "✓ " + msg + "  →  Methoden konvergieren."
            c.font = Font(color="2E7D32", bold=True, size=10)
        row = 5

    if row < 4:
        row = 4
    INPUT_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    REF_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    years = d["years"]
    br = d.get("bridge")

    # --- Cash-flow bridge (FCFF): derive the base FCF from operating metrics, live ---
    # Only when the run used the bridge (Base FCF not typed by hand). The bridge's
    # Base FCF cell then feeds the assumptions block below, so the whole model runs
    # from EBIT down to enterprise value in one formula chain.
    bridge_basefcf_ref = None
    bridge_growth_ref = None
    if br:
        ws.cell(row=row, column=1,
                value="Cash-Flow-Brücke (FCFF) — blaue Zellen ändern").font = Font(bold=True, size=11)
        row += 1
        b_inputs = [
            ("revenue",   f"Revenue {money}", br.get("revenue"), "#,##0"),
            ("ebit",      f"EBIT {money}", br.get("ebit"), "#,##0"),
            ("tax",       "Tax Rate", br.get("tax"), "0.00%"),
            ("da",        f"D&A (= EBITDA − EBIT) {money}", br.get("da"), "#,##0"),
            ("capex_pct", "Capex (% of Revenue)", br.get("capex_pct"), "0.0%"),
            ("nwc_pct",   "ΔNWC (% of Revenue growth)", br.get("nwc_pct"), "0.0%"),
            ("growth",    "FCF Growth", br.get("growth"), "0.0%"),
        ]
        b_row: dict[str, int] = {}
        for key, lbl, val, fmt in b_inputs:
            ws.cell(row=row, column=1, value=lbl)
            c = ws.cell(row=row, column=2, value=val)
            c.number_format = fmt; c.fill = INPUT_FILL
            b_row[key] = row
            row += 1
        REV = f"$B${b_row['revenue']}"; EBIT = f"$B${b_row['ebit']}"
        BTAX = f"$B${b_row['tax']}"; DA = f"$B${b_row['da']}"
        CAPEXP = f"$B${b_row['capex_pct']}"; NWCP = f"$B${b_row['nwc_pct']}"
        bridge_growth_ref = f"$B${b_row['growth']}"

        # derived formula lines. Labels must NOT start with =, + or − : Excel would
        # read such a cell as a formula and show #NAME?. Keep the operator inside the text.
        r_nopat = row
        ws.cell(row=row, column=1, value="NOPAT = EBIT·(1−Tax)")
        ws.cell(row=row, column=2, value=f"={EBIT}*(1-{BTAX})").number_format = "#,##0"; row += 1
        ws.cell(row=row, column=1, value="D&A (non-cash, added back)")
        ws.cell(row=row, column=2, value=f"={DA}").number_format = "#,##0"; row += 1
        r_capex = row
        ws.cell(row=row, column=1, value="Capex = Revenue·Capex%")
        ws.cell(row=row, column=2, value=f"={REV}*{CAPEXP}").number_format = "#,##0"; row += 1
        r_dnwc = row
        ws.cell(row=row, column=1, value="ΔNWC = Revenue·Growth·NWC%")
        ws.cell(row=row, column=2,
                value=f"={REV}*{bridge_growth_ref}*{NWCP}").number_format = "#,##0"; row += 1
        r_bfcf = row
        ws.cell(row=row, column=1,
                value="Base FCF (Jahr 0) = NOPAT + D&A − Capex − ΔNWC").font = Font(bold=True)
        c = ws.cell(row=row, column=2,
                    value=f"=B{r_nopat}+{DA}-B{r_capex}-B{r_dnwc}")
        c.number_format = "#,##0"; c.font = Font(bold=True); c.fill = REF_FILL
        bridge_basefcf_ref = r_bfcf
        row += 2

    # --- Assumptions block (drives the projection) ---
    ws.cell(row=row, column=1,
            value="Annahmen (Eingaben) — blaue Zellen ändern, Modell rechnet neu").font = Font(bold=True, size=11)
    row += 1
    a_row: dict[str, int] = {}
    for key, lbl, val, fmt in [
        ("base_fcf", f"Base FCF (Jahr 0) {money}", d.get("base_fcf"), "#,##0"),
        ("growth",   "FCF Growth", d.get("growth"), "0.0%"),
        ("terminal_growth", "Terminal Growth (g)", d.get("terminal_growth"), "0.0%"),
        ("wacc",     "WACC", d.get("wacc") if d.get("wacc") is not None else wacc, "0.00%"),
    ]:
        ws.cell(row=row, column=1, value=lbl)
        # base_fcf and growth reference the bridge when present (single source of truth);
        # otherwise they are direct inputs (Base FCF typed by hand).
        if br and key == "base_fcf":
            c = ws.cell(row=row, column=2, value=f"=B{bridge_basefcf_ref}"); c.fill = REF_FILL
        elif br and key == "growth":
            c = ws.cell(row=row, column=2, value=f"={bridge_growth_ref}"); c.fill = REF_FILL
        else:
            c = ws.cell(row=row, column=2, value=val); c.fill = INPUT_FILL
        c.number_format = fmt
        a_row[key] = row
        row += 1
    BASEFCF = f"$B${a_row['base_fcf']}"; GROWTH = f"$B${a_row['growth']}"
    TG = f"$B${a_row['terminal_growth']}"; WACC = f"$B${a_row['wacc']}"
    row += 1

    # --- Projection table (columns = years 1..N), all formulas ---
    ws.cell(row=row, column=1, value=f"Cash Flow Projektion {money}").font = Font(bold=True, size=11)
    row += 1
    ws.cell(row=row, column=1, value="Jahr")
    for yr in range(1, years + 1):
        ws.cell(row=row, column=1 + yr, value=f"Jahr {yr}")
    for c in ws[row]:
        if c.value:
            c.fill = HEADER_FILL; c.font = HEADER_FONT
    row += 1

    fcf_row = row
    ws.cell(row=fcf_row, column=1, value="FCF  = Vorjahr·(1+g)")
    for yr in range(1, years + 1):
        col = 1 + yr
        letter = get_column_letter(col)
        if yr == 1:
            f = f"={BASEFCF}*(1+{GROWTH})"
        else:
            prev = get_column_letter(col - 1)
            f = f"={prev}{fcf_row}*(1+{GROWTH})"
        c = ws.cell(row=fcf_row, column=col, value=f); c.number_format = "#,##0"
    row += 1

    df_row = row
    ws.cell(row=df_row, column=1, value="Discount Factor  = 1/(1+WACC)^t")
    for yr in range(1, years + 1):
        c = ws.cell(row=df_row, column=1 + yr, value=f"=1/(1+{WACC})^{yr}")
        c.number_format = "0.000"
    row += 1

    pv_row = row
    ws.cell(row=pv_row, column=1, value="PV of FCF  = FCF·DF")
    for yr in range(1, years + 1):
        letter = get_column_letter(1 + yr)
        c = ws.cell(row=pv_row, column=1 + yr,
                    value=f"={letter}{fcf_row}*{letter}{df_row}")
        c.number_format = "#,##0"
    row += 2

    fcf_last = get_column_letter(1 + years)
    pv_first = get_column_letter(2)
    pv_last = get_column_letter(1 + years)
    df_last = f"{fcf_last}{df_row}"

    # --- EV bridge (all formulas) ---
    ws.cell(row=row, column=1, value="Enterprise Value Bridge").font = Font(bold=True, size=11)
    row += 1
    r_sumpv = row
    ws.cell(row=row, column=1, value="Sum PV of FCFs")
    c = ws.cell(row=row, column=2, value=f"=SUM({pv_first}{pv_row}:{pv_last}{pv_row})")
    c.number_format = "#,##0"; row += 1

    r_tv = row
    ws.cell(row=row, column=1, value="Terminal Value  = FCF_N·(1+g)/(WACC−g)")
    c = ws.cell(row=row, column=2,
                value=f"={fcf_last}{fcf_row}*(1+{TG})/({WACC}-{TG})")
    c.number_format = "#,##0"; row += 1

    r_pvtv = row
    ws.cell(row=row, column=1, value="PV of Terminal Value  = TV·DF_N")
    c = ws.cell(row=row, column=2, value=f"=B{r_tv}*{df_last}")
    c.number_format = "#,##0"; row += 1

    r_ev = row
    ws.cell(row=row, column=1, value="Enterprise Value (DCF)  = ΣPV + PV(TV)").font = Font(bold=True, size=13)
    c = ws.cell(row=row, column=2, value=f"=B{r_sumpv}+B{r_pvtv}")
    c.number_format = "#,##0"; c.font = Font(bold=True, size=13)
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    row += 1

    ws.cell(row=row, column=1, value="TV % of EV  = PV(TV)/EV")
    c = ws.cell(row=row, column=2, value=f"=B{r_pvtv}/B{r_ev}")
    c.number_format = "0.0%"; row += 2

    # --- Sensitivity grid (live: each EV recomputed from the projection cells) ---
    if sens:
        ws.cell(row=row, column=1,
                value="Sensitivität: WACC × Terminal Growth → EV (live)").font = Font(bold=True, size=11)
        row += 1
        # 3×3 grid around the base case, same steps as the app (±1% WACC, ±0.5% g)
        wacc_delta, tg_delta = 0.01, 0.005
        tg_row = row
        ws.cell(row=row, column=1, value="WACC \\ TermG")
        ws.cell(row=row, column=1).fill = HEADER_FILL; ws.cell(row=row, column=1).font = HEADER_FONT
        for j, dtg in enumerate([-tg_delta, 0.0, tg_delta], start=2):
            c = ws.cell(row=row, column=j, value=f"={TG}+({dtg})"); c.number_format = "0.0%"
            c.fill = HEADER_FILL; c.font = HEADER_FONT
        row += 1
        for dwacc in [-wacc_delta, 0.0, wacc_delta]:
            wr = row
            c = ws.cell(row=wr, column=1, value=f"={WACC}+({dwacc})"); c.number_format = "0.00%"
            c.fill = HEADER_FILL; c.font = HEADER_FONT
            for j in range(2, 5):
                tg_col = get_column_letter(j)
                wref = f"$A${wr}"          # this row's WACC
                gref = f"{tg_col}${tg_row}"  # this column's terminal growth
                # EV(w,g) = Σ FCF_t/(1+w)^t  +  [FCF_N·(1+g)/(w−g)]/(1+w)^N
                pv_terms = "+".join(
                    f"{get_column_letter(1+t)}${fcf_row}/(1+{wref})^{t}"
                    for t in range(1, years + 1))
                tv_term = f"({fcf_last}${fcf_row}*(1+{gref})/({wref}-{gref}))/(1+{wref})^{years}"
                c = ws.cell(row=wr, column=j, value=f"={pv_terms}+{tv_term}")
                c.number_format = "#,##0"
            row += 1

    ws.column_dimensions["A"].width = 38
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 15


def _write_precedent_tab(wb, precedent: dict, target_inputs: dict, target_ccy: str):
    """Precedent Transactions tab: deal list, deal-multiple stats, implied valuation.
    Deal multiples embed a control premium → typically higher than trading comps."""
    ws = wb.create_sheet("Precedent Tx")
    money = f"({target_ccy})"
    ws["A1"] = "Precedent Transactions — M&A Deal Comps"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Bewertung auf Basis bezahlter Übernahmepreise. Deal-Multiples enthalten "
                "Kontrollprämie → meist höher als Trading Comps.")
    ws["A2"].font = Font(italic=True, color="666666")

    deals = precedent.get("deals", [])
    stats = precedent.get("stats", {})
    implied = precedent.get("implied", [])

    # --- Deal list ---
    row = 4
    ws.cell(row=row, column=1, value="Deals").font = Font(bold=True, size=11)
    row += 1
    hdr = ["Datum", "Target", "Käufer", "EV", "Revenue", "EBITDA", "EV/Sales", "EV/EBITDA", "EV/EBIT"]
    for i, h in enumerate(hdr, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
    row += 1
    for d in deals:
        ws.cell(row=row, column=1, value=d.date)
        ws.cell(row=row, column=2, value=d.target)
        ws.cell(row=row, column=3, value=d.acquirer)
        for col, val in [(4, d.ev), (5, d.revenue), (6, d.ebitda)]:
            c = ws.cell(row=row, column=col, value=val); c.number_format = "#,##0"
        for col, val in [(7, d.ev_sales), (8, d.ev_ebitda), (9, d.ev_ebit)]:
            c = ws.cell(row=row, column=col, value=round(val, 2) if val else None)
            c.number_format = "0.00x"
        row += 1
    row += 1

    # --- Deal multiple statistics ---
    ws.cell(row=row, column=1, value="Deal Multiple Statistics").font = Font(bold=True, size=11)
    row += 1
    sh = ["Multiple", "N", "Median", "Q1", "Q3", "Min", "Max"]
    for i, h in enumerate(sh, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
    row += 1
    labels = {"ev_sales": "EV / Sales", "ev_ebitda": "EV / EBITDA", "ev_ebit": "EV / EBIT"}
    for key in ["ev_sales", "ev_ebitda", "ev_ebit"]:
        s = stats.get(key)
        if not s:
            continue
        ws.cell(row=row, column=1, value=labels[key])
        ws.cell(row=row, column=2, value=s["n"])
        for col, k in [(3, "median"), (4, "q1"), (5, "q3"), (6, "min"), (7, "max")]:
            c = ws.cell(row=row, column=col, value=round(s[k], 2) if s.get(k) else None)
            c.number_format = "0.00x"
        row += 1
    row += 1

    # --- Implied valuation ---
    ws.cell(row=row, column=1, value=f"Implied Enterprise Value {money}").font = Font(bold=True, size=11)
    row += 1
    ih = ["Multiple", "Operand", "Median", "Low (Q1)", "Base (Median)", "High (Q3)"]
    for i, h in enumerate(ih, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
    row += 1
    for r in implied:
        ws.cell(row=row, column=1, value=r["label"])
        ws.cell(row=row, column=2, value=r["operand"])
        c = ws.cell(row=row, column=3, value=round(r["median"], 2)); c.number_format = "0.00x"
        for col, k in [(4, "low_ev"), (5, "base_ev"), (6, "high_ev")]:
            c = ws.cell(row=row, column=col, value=r.get(k)); c.number_format = "#,##0"
        ws.cell(row=row, column=5).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        row += 1

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    for col in "DEF":
        ws.column_dimensions[col].width = 18   # money columns need room (no ###)
    for col in "GHI":
        ws.column_dimensions[col].width = 12
