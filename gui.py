"""Tkinter GUI for the Peer Group Screening Tool."""
from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path
from tkinter import (
    Tk, Toplevel, Frame, Label, Entry, Text, Button, Checkbutton, StringVar,
    BooleanVar, END, BOTH, LEFT, RIGHT, TOP, X, Y, W, E, NORMAL, DISABLED,
    WORD, FLAT, SOLID, messagebox, Canvas,
)
from tkinter import ttk

import pandas as pd

# --- Color palette (matches Excel output) ---
C_BG       = "#EAEEF3"   # app background
C_CARD     = "#FFFFFF"   # card background
C_NAVY     = "#1F4E78"   # primary / header
C_NAVY2    = "#2F5597"   # accent darker
C_ACCENT   = "#3B6FB5"   # buttons / highlights
C_ACCENT_H = "#2F5C9C"   # button hover
C_TEXT     = "#1E2430"   # main text
C_MUTED    = "#6B7280"   # secondary text
C_LINE     = "#DCE2EA"   # borders
C_GREEN    = "#2E7D32"   # success
C_FIELD    = "#F4F7FB"   # input bg
C_FIELD_F  = "#FFFFFF"   # input bg focused
C_FOCUS    = "#3B6FB5"   # focus ring

from data_source import YFinanceDataSource
from export import export
from multiples import flag_outliers, summary_stats
from screening import apply_filters, rank_by_similarity
from universes import UNIVERSES, list_universes, get_universe
from valuation import compute_implied, summary_range
from wacc import compute_wacc, dcf, dcf_sensitivity
from transactions import load_transactions, deal_multiple_stats, implied_from_transactions
import ai
import discovery
from paths import APP_DIR, OUTPUT_DIR

# Prefix marking an auto-discovery (screened) entry in the sector dropdown
DISCOVERY_PREFIX = "🔍 S&P 500 · "


def all_universe_options():
    """Curated universes + auto-discovery (screened) entries."""
    opts = list(list_universes())
    try:
        for sec in discovery.list_gics_sectors():
            opts.append(f"{DISCOVERY_PREFIX}{sec}")
    except Exception:
        pass
    return opts

OUTPUT_DIR.mkdir(exist_ok=True)


class App:
    def __init__(self, root: Tk):
        self.root = root
        self._last_df = None
        self._last_skipped = []
        root.title("Peer Group Screening Tool")
        root.geometry("880x940")
        root.configure(bg=C_BG)
        root.minsize(760, 600)

        self._setup_styles()

        # ===== Header banner =====
        banner = Frame(root, bg=C_NAVY)
        banner.pack(fill=X, side=TOP)
        inner_b = Frame(banner, bg=C_NAVY)
        inner_b.pack(fill=X, padx=20, pady=(14, 14))
        # Logo mark + title
        logo = Label(inner_b, text="◆", bg=C_NAVY, fg="#7FA8D9",
                     font=("Segoe UI", 22))
        logo.pack(side=LEFT, padx=(0, 12))
        title_box = Frame(inner_b, bg=C_NAVY)
        title_box.pack(side=LEFT, anchor=W)
        Label(title_box, text="Peer Group Screening", bg=C_NAVY, fg="white",
              font=("Segoe UI Semibold", 17)).pack(anchor=W)
        ai_ok = ai.is_available()
        sub = "Trading Comps  ·  DCF  ·  Precedent Transactions  ·  AI Match"
        Label(title_box, text=sub, bg=C_NAVY, fg="#A9C0DC",
              font=("Segoe UI", 9)).pack(anchor=W, pady=(1, 0))
        # AI status pill on the right
        pill_txt = "● AI" if ai_ok else "○ AI"
        pill_col = "#7FE0A0" if ai_ok else "#9FB2C9"
        Label(inner_b, text=pill_txt, bg=C_NAVY, fg=pill_col,
              font=("Segoe UI Semibold", 10)).pack(side=RIGHT, anchor=E)

        # ===== Scrollable body =====
        outer = Frame(root, bg=C_BG)
        outer.pack(fill=BOTH, expand=True)
        canvas = Canvas(outer, bg=C_BG, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body = Frame(canvas, bg=C_BG)
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)
        # Mouse wheel scroll
        def _wheel(e):
            canvas.yview_scroll(int(-e.delta / 120), "units")
        canvas.bind_all("<MouseWheel>", _wheel)

        # ===== Card 1: Setup =====
        c1 = self._card(body, "1 · Setup")
        self.target_name = StringVar(value="My_Peer_Group")
        self.universe_var = StringVar(value=list_universes()[0])
        self.target_ccy = StringVar(value="USD")

        self._row_label(c1, "Target Name")
        Entry(c1, textvariable=self.target_name, **self._entry_kw()).grid(row=0, column=1, columnspan=3, sticky="we", pady=3)

        self._row_label(c1, "Sektor (Universum)", row=1)
        sector_menu = ttk.Combobox(c1, textvariable=self.universe_var,
                                   values=all_universe_options(), state="readonly")
        sector_menu.grid(row=1, column=1, columnspan=3, sticky="we", pady=3)
        sector_menu.bind("<<ComboboxSelected>>", self.on_universe_change)

        self._row_label(c1, "Zielwährung", row=2)
        ccy_menu = ttk.Combobox(c1, textvariable=self.target_ccy,
                                values=["USD", "EUR", "GBP", "CHF"], state="readonly", width=8)
        ccy_menu.grid(row=2, column=1, sticky=W, pady=3)
        ccy_menu.bind("<<ComboboxSelected>>", self.on_ccy_change)
        Label(c1, text="alle Beträge werden umgerechnet", bg=C_CARD, fg=C_MUTED,
              font=("Segoe UI", 8)).grid(row=2, column=2, columnspan=2, sticky=W)
        self.make_pptx = BooleanVar(value=False)
        Checkbutton(c1, text="PowerPoint-Deck zusätzlich erstellen (.pptx)",
                    variable=self.make_pptx, bg=C_CARD, activebackground=C_CARD,
                    fg=C_TEXT, font=("Segoe UI", 9), highlightthickness=0).grid(
            row=3, column=0, columnspan=4, sticky=W, pady=(4, 0))
        c1.columnconfigure(1, weight=1)
        c1.columnconfigure(3, weight=1)

        # ===== Card 2: Filter =====
        c2 = self._card(body, "2 · Filter (Umsatz / Marge — leer = aus)")
        self.filter_label = self._mini_label(c2, "Umsatz in Mio USD", 0)
        self.rev_min = self._field(c2, "Revenue min", "200", 1, 0)
        self.rev_max = self._field(c2, "Revenue max", "2000", 1, 2)
        self.mar_min = self._field(c2, "EBITDA-Marge min %", "", 2, 0)
        self.mar_max = self._field(c2, "EBITDA-Marge max %", "", 2, 2)

        # ===== Card 3: Target Financials =====
        c3 = self._card(body, "3 · Target-Financials (für Implied Valuation)")
        self.target_fin_label = self._mini_label(c3, "in Mio USD", 0)
        self.t_revenue    = self._field(c3, "Revenue", "", 1, 0)
        self.t_ebitda     = self._field(c3, "EBITDA", "", 1, 2)
        self.t_ebit       = self._field(c3, "EBIT", "", 1, 4)
        self.t_net_income = self._field(c3, "Net Income", "", 2, 0)
        self.t_net_debt   = self._field(c3, "Net Debt", "", 2, 2)
        # Precedent Transactions (optional, 4th valuation method) — pluggable source
        Label(c3, text="Precedent Transactions (optional — M&A-Deal-Multiples):",
              bg=C_CARD, fg=C_MUTED, font=("Segoe UI", 8)).grid(row=3, column=0, columnspan=6, sticky=W, pady=(6, 0))
        Label(c3, text="Deal-Quelle", bg=C_CARD, fg=C_TEXT, anchor=W,
              font=("Segoe UI", 9)).grid(row=4, column=0, sticky=W, padx=(0, 4), pady=3)
        self.tx_source = StringVar(value="csv")
        tx_menu = ttk.Combobox(c3, textvariable=self.tx_source, state="readonly", width=22,
                               values=["csv  (lokale Datei)", "fmp  (FMP API — Live-Feed)"])
        tx_menu.set("csv  (lokale Datei)")
        tx_menu.grid(row=4, column=1, columnspan=2, sticky=W, pady=3)
        self.tx_csv = self._field(c3, "CSV / Key", "transactions.csv", 4, 4)

        # ===== Card 4: WACC / DCF =====
        c4 = self._card(body, "4 · WACC / DCF Annahmen (frei änderbar)")
        self._mini_label(c4, "% als Zahl eingeben, z.B. 2.5", 0)
        self.w_rf        = self._field(c4, "Risk-free %", "2.5", 1, 0)
        self.w_mrp       = self._field(c4, "Market Prem %", "5.5", 1, 2)
        self.w_tax       = self._field(c4, "Tax %", "25", 1, 4)
        self.w_kd        = self._field(c4, "Cost of Debt %", "5.0", 1, 6)
        self.w_target_de = self._field(c4, "Target D/E", "0.3", 2, 0)
        self.w_size_prem = self._field(c4, "Size-Premium %", "0", 2, 2)
        self.d_base_fcf  = self._field(c4, "Base FCF (Mio)", "", 2, 4)
        self.d_growth    = self._field(c4, "FCF-Wachstum %", "8", 2, 6)
        self.d_term_g    = self._field(c4, "Term. Growth %", "2.5", 3, 0)

        # ===== Card 5: AI Business Match =====
        c5 = self._card(body, "5 · Target-Geschäftsmodell (AI Match Score)")
        badge_txt = "● Gemini AI verbunden" if ai_ok else "● Kein API-Key (.env)"
        badge_col = C_GREEN if ai_ok else "#C0392B"
        Label(c5, text=badge_txt, bg=C_CARD, fg=badge_col,
              font=("Segoe UI", 8, "bold")).grid(row=0, column=0, columnspan=8, sticky=W)
        Label(c5, text="1-3 Sätze: Was, Kunden, Vertrieb, Region, Erlösmodell.",
              bg=C_CARD, fg=C_MUTED, font=("Segoe UI", 8)).grid(row=1, column=0, columnspan=8, sticky=W, pady=(0, 3))
        self.target_desc = Text(c5, height=3, wrap=WORD, font=("Segoe UI", 9),
                                relief=SOLID, bd=1, bg=C_FIELD, highlightthickness=0)
        self.target_desc.grid(row=2, column=0, columnspan=8, sticky="we")
        c5.columnconfigure(0, weight=1)

        # ===== Card 6: Ticker Universe =====
        c6 = self._card(body, "6 · Ticker-Universum (editierbar)")
        Label(c6, text="Vorausgefüllt nach Sektor — Ticker ergänzen oder löschen.",
              bg=C_CARD, fg=C_MUTED, font=("Segoe UI", 8)).grid(row=0, column=0, sticky=W)
        self.tickers_text = Text(c6, height=7, wrap=WORD, font=("Consolas", 9),
                                 relief=FLAT, bd=0, bg=C_FIELD, highlightthickness=1,
                                 highlightbackground=C_LINE, highlightcolor=C_FOCUS,
                                 padx=6, pady=5)
        self.tickers_text.grid(row=1, column=0, sticky="we")
        c6.columnconfigure(0, weight=1)
        self.on_universe_change()

        # ===== Run + Shortlist buttons =====
        runwrap = Frame(body, bg=C_BG)
        runwrap.pack(fill=X, padx=16, pady=(12, 12))
        self.run_btn = Button(
            runwrap, text="▶   PEER GROUP ANALYSIEREN",
            command=self.on_run, height=2, relief=FLAT, cursor="hand2", bd=0,
            bg=C_ACCENT, fg="white", activebackground=C_ACCENT_H, activeforeground="white",
            font=("Segoe UI Semibold", 12),
        )
        self.run_btn.pack(side=LEFT, fill=X, expand=True)
        self._hover(self.run_btn, C_ACCENT, C_ACCENT_H)
        self.shortlist_btn = Button(
            runwrap, text="☑   Short List", state=DISABLED,
            command=self.open_shortlist, height=2, relief=FLAT, cursor="hand2", bd=0,
            bg=C_NAVY, fg="white", activebackground=C_NAVY2, activeforeground="white",
            disabledforeground="#94A3B8", font=("Segoe UI Semibold", 11),
        )
        self.shortlist_btn.pack(side=LEFT, fill=X, padx=(10, 0), ipadx=10)
        self._hover_guarded(self.shortlist_btn, C_NAVY, C_NAVY2)

        # ===== Status log =====
        logcard = self._card(body, "Status")
        self.log = Text(logcard, height=11, bg="#161C26", fg="#CFE0F0", state=DISABLED,
                        wrap=WORD, font=("Consolas", 9), relief=FLAT, highlightthickness=0,
                        insertbackground="white", padx=10, pady=8)
        self.log.grid(row=0, column=0, sticky="we")
        logcard.columnconfigure(0, weight=1)

        # Restore last session (after all widgets built)
        self._load_settings()
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---- settings persistence ----

    SETTINGS_FILE = APP_DIR / "settings.json"

    def _collect_settings(self) -> dict:
        return {
            "target_name": self.target_name.get(),
            "universe": self.universe_var.get(),
            "currency": self.target_ccy.get(),
            "rev_min": self.rev_min.get(), "rev_max": self.rev_max.get(),
            "mar_min": self.mar_min.get(), "mar_max": self.mar_max.get(),
            "t_revenue": self.t_revenue.get(), "t_ebitda": self.t_ebitda.get(),
            "t_ebit": self.t_ebit.get(), "t_net_income": self.t_net_income.get(),
            "t_net_debt": self.t_net_debt.get(),
            "w_rf": self.w_rf.get(), "w_mrp": self.w_mrp.get(), "w_tax": self.w_tax.get(),
            "w_kd": self.w_kd.get(), "w_target_de": self.w_target_de.get(),
            "w_size_prem": self.w_size_prem.get(),
            "d_base_fcf": self.d_base_fcf.get(), "d_growth": self.d_growth.get(),
            "d_term_g": self.d_term_g.get(),
            "tx_csv": self.tx_csv.get(),
            "tx_source": self.tx_source.get(),
            "make_pptx": self.make_pptx.get(),
            "target_desc": self.target_desc.get("1.0", END).strip(),
            "tickers": self.tickers_text.get("1.0", END).strip(),
        }

    def _save_settings(self):
        import json
        try:
            self.SETTINGS_FILE.write_text(
                json.dumps(self._collect_settings(), ensure_ascii=False, indent=2),
                encoding="utf-8")
        except Exception:
            pass

    def _load_settings(self):
        import json
        if not self.SETTINGS_FILE.exists():
            return
        try:
            s = json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        # Set universe first WITHOUT triggering the auto-fill (which would overwrite tickers)
        if s.get("universe") in all_universe_options():
            self.universe_var.set(s["universe"])
        setters = {
            "target_name": self.target_name, "currency": self.target_ccy,
            "rev_min": self.rev_min, "rev_max": self.rev_max,
            "mar_min": self.mar_min, "mar_max": self.mar_max,
            "t_revenue": self.t_revenue, "t_ebitda": self.t_ebitda, "t_ebit": self.t_ebit,
            "t_net_income": self.t_net_income, "t_net_debt": self.t_net_debt,
            "w_rf": self.w_rf, "w_mrp": self.w_mrp, "w_tax": self.w_tax, "w_kd": self.w_kd,
            "w_target_de": self.w_target_de, "w_size_prem": self.w_size_prem,
            "d_base_fcf": self.d_base_fcf, "d_growth": self.d_growth, "d_term_g": self.d_term_g,
            "tx_csv": self.tx_csv,
        }
        for key, var in setters.items():
            if s.get(key) is not None:
                var.set(s[key])
        if s.get("make_pptx") is not None:
            self.make_pptx.set(bool(s["make_pptx"]))
        if s.get("tx_source"):
            self.tx_source.set(s["tx_source"])
        if s.get("target_desc"):
            self.target_desc.delete("1.0", END)
            self.target_desc.insert(END, s["target_desc"])
        if s.get("tickers"):
            self.tickers_text.delete("1.0", END)
            self.tickers_text.insert(END, s["tickers"])
        self.on_ccy_change()
        self.write_log("↻ Letzter Suchauftrag wiederhergestellt.")

    def _on_close(self):
        self._save_settings()
        self.root.destroy()

    # ---- style helpers ----

    def _setup_styles(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TCombobox", fieldbackground=C_FIELD, background="white",
                     bordercolor=C_LINE, arrowcolor=C_NAVY, padding=5, relief="flat")
        st.map("TCombobox",
               fieldbackground=[("readonly", C_FIELD)],
               selectbackground=[("readonly", C_FIELD)],
               selectforeground=[("readonly", C_TEXT)],
               bordercolor=[("focus", C_FOCUS)])
        st.configure("Vertical.TScrollbar", background="#C7D0DB", troughcolor=C_BG,
                     bordercolor=C_BG, arrowcolor=C_NAVY, relief="flat", width=12)
        st.map("Vertical.TScrollbar", background=[("active", C_ACCENT)])

    def _hover(self, widget, normal, hover):
        widget.bind("<Enter>", lambda e: widget.config(bg=hover))
        widget.bind("<Leave>", lambda e: widget.config(bg=normal))

    def _hover_guarded(self, widget, normal, hover):
        """Hover that respects disabled state."""
        widget.bind("<Enter>", lambda e: widget.config(bg=hover) if str(widget["state"]) != "disabled" else None)
        widget.bind("<Leave>", lambda e: widget.config(bg=normal) if str(widget["state"]) != "disabled" else None)

    def _focus_ring(self, entry):
        """Highlight an entry while focused (focus ring + white bg)."""
        entry.config(highlightthickness=1, highlightbackground=C_LINE, highlightcolor=C_FOCUS)
        entry.bind("<FocusIn>", lambda e: entry.config(bg=C_FIELD_F))
        entry.bind("<FocusOut>", lambda e: entry.config(bg=C_FIELD))

    def _card(self, parent, title):
        """Creates a titled white card with an accent number badge; returns inner grid frame."""
        wrap = Frame(parent, bg=C_BG)
        wrap.pack(fill=X, padx=16, pady=(9, 0))

        # Title row with accent bar
        trow = Frame(wrap, bg=C_BG)
        trow.pack(fill=X, pady=(0, 3))
        bar = Frame(trow, bg=C_ACCENT, width=3, height=14)
        bar.pack(side=LEFT, padx=(0, 7)); bar.pack_propagate(False)
        Label(trow, text=title.upper(), bg=C_BG, fg=C_NAVY,
              font=("Segoe UI Semibold", 9)).pack(side=LEFT, anchor=W)

        # Card body with subtle border
        card = Frame(wrap, bg=C_CARD, highlightbackground=C_LINE,
                     highlightthickness=1, bd=0)
        card.pack(fill=X)
        inner = Frame(card, bg=C_CARD)
        inner.pack(fill=X, padx=14, pady=12)
        return inner

    def _entry_kw(self):
        return dict(relief=FLAT, bd=0, bg=C_FIELD, font=("Segoe UI", 10),
                    highlightthickness=1, highlightbackground=C_LINE,
                    highlightcolor=C_FOCUS, insertbackground=C_TEXT)

    def _row_label(self, parent, text, row=0):
        Label(parent, text=text, bg=C_CARD, fg=C_TEXT, anchor=W,
              font=("Segoe UI", 10)).grid(row=row, column=0, sticky=W, padx=(0, 8), pady=4)

    def _mini_label(self, parent, text, row):
        lbl = Label(parent, text=text, bg=C_CARD, fg=C_MUTED, anchor=W,
                    font=("Segoe UI", 8))
        lbl.grid(row=row, column=0, columnspan=8, sticky=W, pady=(0, 5))
        return lbl

    def _field(self, parent, label, default, row, col):
        """Compact labelled entry in a grid. Returns the StringVar."""
        Label(parent, text=label, bg=C_CARD, fg=C_MUTED, anchor=W,
              font=("Segoe UI", 9)).grid(row=row, column=col, sticky=W, padx=(0, 5), pady=4)
        var = StringVar(value=default)
        e = Entry(parent, textvariable=var, width=11, **self._entry_kw())
        e.grid(row=row, column=col + 1, sticky=W, padx=(0, 16), pady=4, ipady=3)
        self._focus_ring(e)
        return var

    # ---- helpers ----

    def on_universe_change(self, event=None):
        name = self.universe_var.get()
        if name.startswith(DISCOVERY_PREFIX):
            # Auto-discovery: screen S&P 500 constituents by GICS sector
            sector = name[len(DISCOVERY_PREFIX):]
            tickers = discovery.screen(sector=sector)
            clean = "SP500_" + sector.replace(" ", "_")
        else:
            tickers = get_universe(name)
            clean = name.split("—")[0].strip().replace(" / ", "_").replace(" ", "_")
        self.tickers_text.delete("1.0", END)
        self.tickers_text.insert(END, ", ".join(tickers))
        self.target_name.set(clean)

    def on_ccy_change(self, event=None):
        ccy = self.target_ccy.get()
        self.filter_label.config(text=f"Umsatz in Mio {ccy}")
        self.target_fin_label.config(text=f"in Mio {ccy}")

    def parse_tickers(self) -> list[str]:
        raw = self.tickers_text.get("1.0", END)
        parts = [p.strip().upper() for p in raw.replace("\n", ",").split(",")]
        # dedupe while preserving order
        seen, out = set(), []
        for p in parts:
            if p and p not in seen:
                seen.add(p); out.append(p)
        return out

    def parse_float_mio(self, s: str):
        s = (s or "").strip().replace(",", ".")
        if not s:
            return None
        try:
            return float(s) * 1_000_000
        except ValueError:
            return None

    def parse_pct(self, s: str):
        s = (s or "").strip().replace(",", ".").rstrip("%")
        if not s:
            return None
        try:
            return float(s) / 100.0
        except ValueError:
            return None

    def write_log(self, msg: str):
        self.log.config(state=NORMAL)
        self.log.insert(END, msg + "\n")
        self.log.see(END)
        self.log.config(state=DISABLED)
        self.root.update_idletasks()

    # ---- main action ----

    def on_run(self):
        tickers = self.parse_tickers()
        if not tickers:
            messagebox.showerror("Fehler", "Keine Ticker im Universum. Sektor auswählen.")
            return
        self._save_settings()  # persist this search before running
        self.run_btn.config(state=DISABLED, text="Läuft …")
        self.log.config(state=NORMAL); self.log.delete("1.0", END); self.log.config(state=DISABLED)
        t = threading.Thread(target=self._run_worker, args=(tickers,), daemon=True)
        t.start()

    def open_shortlist(self):
        """Opens a checkbox window to pick the final peer set, then recomputes."""
        if self._last_df is None or self._last_df.empty:
            messagebox.showinfo("Short List", "Erst eine Analyse laufen lassen.")
            return
        df = self._last_df

        win = Toplevel(self.root)
        win.title("Short List — Peers auswählen")
        win.geometry("640x600")
        win.configure(bg=C_BG)

        Label(win, text="Short List zusammenstellen", bg=C_NAVY, fg="white",
              font=("Segoe UI Semibold", 13)).pack(fill=X, ipady=8)
        Label(win, text="Häkchen = fließt in Statistik/Valuation. Vorausgewählt: Match-Score ≥7 (sonst alle im Filter).",
              bg=C_BG, fg=C_MUTED, font=("Segoe UI", 8), wraplength=600, justify=LEFT).pack(fill=X, padx=12, pady=(6, 4))

        # Scrollable checkbox list
        outer = Frame(win, bg=C_BG)
        outer.pack(fill=BOTH, expand=True, padx=12)
        cv = Canvas(outer, bg=C_CARD, highlightthickness=1, highlightbackground=C_LINE)
        sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        lst = Frame(cv, bg=C_CARD)
        lst.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=lst, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side=LEFT, fill=BOTH, expand=True)
        sb.pack(side=RIGHT, fill=Y)
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units"))

        # Determine default selection
        has_scores = "match_score" in df.columns and df["match_score"].notna().any()
        vars_by_symbol = {}

        # Header row
        hdr = Frame(lst, bg=C_CARD)
        hdr.pack(fill=X, pady=(4, 2))
        Label(hdr, text="  ✓   Ticker        Score   Company", bg=C_CARD, fg=C_NAVY,
              font=("Consolas", 9, "bold"), anchor=W).pack(fill=X, padx=6)

        for _, r in df.iterrows():
            sym = r["symbol"]
            score = r.get("match_score")
            included = bool(r.get("included"))
            outlier = bool(r.get("is_outlier"))
            # Default: score>=7 if scores exist, else included & not outlier
            if has_scores and pd.notna(score):
                default = int(score) >= 7
            else:
                default = included and not outlier
            v = BooleanVar(value=default)
            vars_by_symbol[sym] = v
            score_str = f"{int(score)}" if pd.notna(score) else "-"
            flag = ""
            if not included:
                flag = " (gefiltert)"
            elif outlier:
                flag = " (Ausreißer)"
            name = (r.get("name") or "")[:32]
            row = Frame(lst, bg=C_CARD)
            row.pack(fill=X)
            Checkbutton(
                row, variable=v, bg=C_CARD, activebackground=C_CARD,
                highlightthickness=0,
                text=f"{sym:<12} [{score_str:>2}]  {name}{flag}",
                font=("Consolas", 9), anchor=W,
            ).pack(fill=X, padx=6, anchor=W)

        # Buttons
        btnbar = Frame(win, bg=C_BG)
        btnbar.pack(fill=X, padx=12, pady=10)

        def _set_all(val):
            for v in vars_by_symbol.values():
                v.set(val)

        Button(btnbar, text="Alle", command=lambda: _set_all(True),
               relief=FLAT, bg=C_LINE, cursor="hand2").pack(side=LEFT)
        Button(btnbar, text="Keine", command=lambda: _set_all(False),
               relief=FLAT, bg=C_LINE, cursor="hand2", padx=6).pack(side=LEFT, padx=6)

        def _recompute():
            selected = [s for s, v in vars_by_symbol.items() if v.get()]
            if len(selected) < 2:
                messagebox.showwarning("Short List", "Mindestens 2 Peers wählen.")
                return
            win.destroy()
            self.run_btn.config(state=DISABLED)
            self.shortlist_btn.config(state=DISABLED)
            t = threading.Thread(target=self._shortlist_worker, args=(selected,), daemon=True)
            t.start()

        Button(btnbar, text="▶  Short List berechnen", command=_recompute,
               relief=FLAT, bg=C_ACCENT, fg="white", cursor="hand2",
               font=("Segoe UI Semibold", 10), padx=12).pack(side=RIGHT)

    def _shortlist_worker(self, selected_symbols):
        try:
            self.log.config(state=NORMAL); self.log.delete("1.0", END); self.log.config(state=DISABLED)
            self.write_log(f"[Short List] {len(selected_symbols)} Peers ausgewählt.")
            df = self._last_df.copy()
            # Keep only selected; force them included, drop outlier flag so all selected count
            df = df[df["symbol"].isin(selected_symbols)].copy()
            df["included"] = True
            df["excluded_reason"] = ""
            df["is_outlier"] = False
            df["outlier_flags"] = ""
            self.write_log("[Short List] Neuberechnung mit ausgewählten Peers …")
            self._compute_and_export(df, len(selected_symbols), self._last_skipped, suffix="SHORTLIST")
        except Exception as e:
            import traceback
            self.write_log("FEHLER:\n" + traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.run_btn.config(state=NORMAL, text="▶  Peer Group analysieren"))
            self.root.after(0, lambda: self.shortlist_btn.config(state=NORMAL))

    def _fmp_key(self) -> str:
        """Read FMP API key from config.json or env, if present."""
        import os, json
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

    def _compute_and_export(self, df, tickers_count, skipped, suffix=""):
        """Shared compute+export. Reads GUI assumption fields. Used by full run AND shortlist.
        df must already have included/is_outlier/match columns. suffix tags the filename.
        """
        target_ccy = self.target_ccy.get()
        target_name = (self.target_name.get() or "peer_group").strip().replace(" ", "_")

        summary = summary_stats(df[df["included"]], exclude_outliers=True)

        # ---- Implied Valuation ----
        target_inputs = {
            "revenue":    self.parse_float_mio(self.t_revenue.get()),
            "ebitda":     self.parse_float_mio(self.t_ebitda.get()),
            "ebit":       self.parse_float_mio(self.t_ebit.get()),
            "net_income": self.parse_float_mio(self.t_net_income.get()),
            "net_debt":   self.parse_float_mio(self.t_net_debt.get()),
        }
        has_target = any(v is not None for v in target_inputs.values())
        implied = compute_implied(target_inputs, summary) if has_target else []
        if has_target and implied:
            self.write_log(f"      Implied Valuation: {len(implied)} Multiples")

        # ---- WACC + DCF ----
        wacc_result, dcf_result, dcf_sens = {}, {}, []
        try:
            rf = self.parse_pct(self.w_rf.get())
            mrp = self.parse_pct(self.w_mrp.get())
            tax = self.parse_pct(self.w_tax.get())
            kd = self.parse_pct(self.w_kd.get())
            de_raw = (self.w_target_de.get() or "").strip().replace(",", ".")
            target_de = float(de_raw) if de_raw else None
            size_prem = self.parse_pct(self.w_size_prem.get()) or 0.0
            beta_pool = df[df["included"]]
            if "match_score" in df.columns and df["match_score"].notna().any():
                matched = beta_pool[pd.to_numeric(beta_pool["match_score"], errors="coerce") >= 6]
                if len(matched) >= 3:
                    beta_pool = matched
            peer_betas = beta_pool[["symbol", "name", "beta", "market_cap", "net_debt"]].to_dict(orient="records")
            if all(v is not None for v in (rf, mrp, tax, kd)) and target_de is not None:
                wacc_result = compute_wacc(peer_betas, rf, mrp, tax, kd, target_de, size_premium=size_prem)
                if wacc_result:
                    self.write_log(f"      WACC: {wacc_result['wacc']:.2%} (Beta {wacc_result['relevered_beta']:.2f})")
                    # D/E outlier diagnostic
                    if wacc_result.get("n_de_outliers"):
                        wc = wacc_result.get("wacc_clean")
                        wc_str = f", bereinigt ~{wc:.2%}" if wc else ""
                        self.write_log(f"      ⚠ {wacc_result['n_de_outliers']} Peer(s) mit D/E >200% — Beta evtl. verzerrt{wc_str}")
                    base_fcf = self.parse_float_mio(self.d_base_fcf.get())
                    growth = self.parse_pct(self.d_growth.get())
                    term_g = self.parse_pct(self.d_term_g.get())
                    if base_fcf and growth is not None and term_g is not None:
                        dcf_result = dcf(base_fcf, growth, term_g, wacc_result["wacc"])
                        dcf_sens = dcf_sensitivity(base_fcf, growth, term_g, wacc_result["wacc"])
                        if dcf_result.get("enterprise_value"):
                            self.write_log(f"      DCF EV: {dcf_result['enterprise_value']/1e6:,.0f} Mio {target_ccy}")
        except Exception as e:
            self.write_log(f"      WACC/DCF-Fehler: {e}")

        # ---- Convergence check: DCF vs Comps median EV ----
        convergence = None
        try:
            if dcf_result and dcf_result.get("enterprise_value") and implied:
                comps_evs = [r["base_ev"] for r in implied if r.get("base_ev")]
                if comps_evs:
                    comps_median = sorted(comps_evs)[len(comps_evs) // 2]
                    dcf_ev = dcf_result["enterprise_value"]
                    gap = abs(dcf_ev - comps_median) / comps_median
                    convergence = {"comps_median_ev": comps_median, "dcf_ev": dcf_ev, "gap": gap}
                    if gap > 0.5:
                        self.write_log(f"      ⚠ DCF/Comps divergieren {gap:.0%} — WACC/Annahmen prüfen")
        except Exception:
            pass

        # ---- Precedent Transactions (4th method) — pluggable source ----
        precedent = None
        try:
            source = "fmp" if self.tx_source.get().startswith("fmp") else "csv"
            field_val = (self.tx_csv.get() or "").strip()
            if has_target and (field_val or source == "fmp"):
                universe = self.universe_var.get()
                if source == "csv":
                    csv_path = APP_DIR / field_val if not Path(field_val).is_absolute() else Path(field_val)
                    all_deals = load_transactions(csv_path, scale=1e6, source="csv")
                    deals = load_transactions(csv_path, scale=1e6, source="csv", universe_name=universe)
                else:  # fmp — field holds the API key (fallback to config key)
                    key = field_val or self._fmp_key()
                    self.write_log("      Deal-Quelle: FMP API (Live-Feed) …")
                    all_deals = load_transactions(source="fmp", api_key=key)
                    deals = all_deals  # feed has no sector tag → no filter
                    n_valued = sum(1 for d in deals if d.ev)
                    self.write_log(f"      FMP: {len(all_deals)} Deals geladen, {n_valued} mit Deal-Wert.")
                    if n_valued == 0:
                        self.write_log("      ⚠ FMP-Free liefert keine Deal-Werte → keine Multiples. (Adapter funktioniert, Daten fehlen.)")

                if deals:
                    stats = deal_multiple_stats(deals)
                    tx_implied = implied_from_transactions(target_inputs, stats)
                    if tx_implied:
                        precedent = {"deals": deals, "stats": stats, "implied": tx_implied}
                        if source == "csv":
                            self.write_log(f"      Precedent Transactions: {len(all_deals)} Deals in Datei, "
                                           f"{len(deals)} passen zu '{universe.split('—')[0].strip()}'")
                    elif source == "csv" and all_deals:
                        self.write_log(f"      {len(all_deals)} Deals in Datei, aber keine passen zum Sektor — kein Precedent-Balken.")
                elif source == "csv":
                    self.write_log(f"      Keine Deals in {field_val} gefunden.")
        except Exception as e:
            self.write_log(f"      Transactions-Fehler: {e}")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        tag = f"_{suffix}" if suffix else ""
        # One folder per run — keeps Excel + PowerPoint together, no clutter
        run_dir = OUTPUT_DIR / f"{target_name}{tag}_{ts}"
        run_dir.mkdir(parents=True, exist_ok=True)
        self._run_dir = run_dir
        out_path = run_dir / f"{target_name}{tag}.xlsx"

        target_meta = {
            "Target": target_name + (f"  [{suffix}]" if suffix else ""),
            "Sektor / Universum": self.universe_var.get(),
            "Zielwährung": target_ccy,
            "Ticker bewertet": tickers_count,
            "Daten erhalten": len(df),
            "Übersprungen (keine Daten)": len(skipped),
            "Eingeschlossen (passender Filter)": int(df["included"].sum()),
            "Ausreißer markiert": int(df["is_outlier"].sum()),
            f"Revenue Range (Mio {target_ccy})": f"{self.rev_min.get() or '–'} – {self.rev_max.get() or '–'}",
            "EBITDA-Marge Range": f"{self.mar_min.get() or '–'}% – {self.mar_max.get() or '–'}%",
            "Run timestamp": datetime.now().isoformat(timespec="seconds"),
        }
        if suffix == "SHORTLIST":
            target_meta["Modus"] = "Short List (manuell ausgewählte Peers)"

        self.write_log("      Schreibe Excel …")
        export(df, summary, target_meta, out_path, skipped=skipped,
               implied=implied, target_inputs=target_inputs if has_target else None,
               wacc_result=wacc_result or None, dcf_result=dcf_result or None,
               dcf_sens=dcf_sens or None, convergence=convergence, precedent=precedent)
        self.write_log(f"✅ Fertig: {out_path}")
        n_incl = int((df["included"] & ~df["is_outlier"]).sum())
        self.write_log(f"   Peers in Statistik: {n_incl}")

        # ---- Optional PowerPoint deck ----
        if self.make_pptx.get() and has_target:
            try:
                try:
                    from pptx_export import build_pptx
                except ModuleNotFoundError:
                    import sys, subprocess
                    self.write_log("      python-pptx fehlt — installiere einmalig …")
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "python-pptx"])
                    import site; site.main()
                    from pptx_export import build_pptx
                self.write_log("      Erstelle PowerPoint-Deck …")
                pptx_path = out_path.with_suffix(".pptx")
                build_pptx(
                    pptx_path,
                    target_name=self.target_name.get() or "Target",
                    sector=self.universe_var.get(),
                    ccy=target_ccy,
                    summary_df=summary,
                    implied=implied,
                    wacc_result=wacc_result or None,
                    dcf_result=dcf_result or None,
                    precedent=precedent,
                    n_peers_included=n_incl,
                )
                self.write_log(f"   📊 PowerPoint: {pptx_path}")
            except Exception as e:
                self.write_log(f"      PowerPoint-Fehler: {e}")

        # Open the run folder (contains Excel + PowerPoint together)
        self.write_log(f"📁 Ordner: {run_dir}")
        try:
            os.startfile(str(run_dir))
        except Exception:
            try:
                os.startfile(str(out_path))
            except Exception:
                pass

    def _run_worker(self, tickers: list[str]):
        try:
            target_name = (self.target_name.get() or "peer_group").strip().replace(" ", "_")
            rev_min = self.parse_float_mio(self.rev_min.get())
            rev_max = self.parse_float_mio(self.rev_max.get())
            mar_min = self.parse_pct(self.mar_min.get())
            mar_max = self.parse_pct(self.mar_max.get())

            target_ccy = self.target_ccy.get()
            self.write_log(f"[1/4] Universum: {len(tickers)} Ticker zu prüfen. Zielwährung: {target_ccy}")
            self.write_log(f"      Filter: Revenue {self.rev_min.get() or '–'}–{self.rev_max.get() or '–'} Mio {target_ccy}")
            ds = YFinanceDataSource(target_currency=target_ccy)
            rows = []
            skipped: list[tuple[str, str]] = []  # (ticker, reason)
            self.write_log("[2/4] Hole Financials von Yahoo Finance …")
            for i, sym in enumerate(tickers, 1):
                try:
                    fin = ds.fetch(sym)
                except Exception as e:
                    self.write_log(f"   [{i:>3}/{len(tickers)}] {sym}: Fehler — {e}")
                    skipped.append((sym, f"error: {e}"))
                    continue
                if fin is None:
                    self.write_log(f"   [{i:>3}/{len(tickers)}] {sym}: übersprungen (keine Equity-Daten)")
                    skipped.append((sym, "no equity data / wrong quote type"))
                    continue
                rows.append(fin.__dict__)
                rev_str = f"{fin.revenue/1e6:>7.0f}M" if fin.revenue else "    n/a"
                self.write_log(f"   [{i:>3}/{len(tickers)}] {sym:<12} {(fin.name or '')[:30]:<30} rev={rev_str}")

            if not rows:
                self.write_log("Keine Daten erhalten — Abbruch.")
                messagebox.showerror("Fehler", "Keine Daten erhalten.")
                return

            df = pd.DataFrame(rows)

            # ---- AI Business Model Match ----
            target_desc = self.target_desc.get("1.0", END).strip()
            if target_desc and ai.is_available():
                self.write_log(f"      AI bewertet Business-Model-Match (Gemini, 1 Call für {len(df)} Peers)…")
                try:
                    peer_inputs = df[["symbol", "name", "sector", "industry", "business_summary"]].to_dict(orient="records")
                    scores = ai.score_peers(target_desc, peer_inputs)
                    if scores:
                        df["match_score"] = df["symbol"].map(lambda s: scores.get(s, {}).get("score"))
                        df["match_rationale"] = df["symbol"].map(lambda s: scores.get(s, {}).get("rationale"))
                        df["business_profile"] = df["symbol"].map(lambda s: scores.get(s, {}).get("business_profile"))
                        self.write_log(f"      AI-Match-Scores: {sum(1 for v in scores.values() if v.get('score'))} Peers bewertet")
                    else:
                        self.write_log("      AI gab keine verwertbare Antwort.")
                except Exception as e:
                    self.write_log(f"      AI-Fehler: {e}")

            self.write_log("[3/4] Filter + Ausreißererkennung …")
            df = apply_filters(df, rev_min, rev_max, mar_min, mar_max)
            df = flag_outliers(df)
            rev_mid = ((rev_min or 0) + (rev_max or 0)) / 2 or None
            # If we have AI scores, sort included peers by score (desc), then by size proximity
            if "match_score" in df.columns and df["match_score"].notna().any():
                df["_score_sort"] = df["match_score"].fillna(-1)
                df = df.sort_values(
                    by=["included", "_score_sort"], ascending=[False, False], na_position="last"
                ).drop(columns=["_score_sort"]).reset_index(drop=True)
            else:
                df = rank_by_similarity(df, target_revenue=rev_mid)
            # Store fetched+scored df so Shortlist can recompute without re-fetching
            self._last_df = df.copy()
            self._last_skipped = skipped

            self._compute_and_export(df, len(tickers), skipped, suffix="")
            # Enable shortlist button on the UI thread
            self.root.after(0, lambda: self.shortlist_btn.config(state=NORMAL))

        except Exception as e:
            import traceback
            self.write_log("FEHLER:\n" + traceback.format_exc())
            messagebox.showerror("Fehler", str(e))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state=NORMAL, text="▶  Peer Group analysieren"))


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
