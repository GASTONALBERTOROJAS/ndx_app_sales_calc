"""
NDX Sales Calculation Tool - GUI Application (tkinter).

Corporate-styled interface with Nordex branding.
"""

import sys
import queue
import threading
import pandas as pd
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

from core.extractor import run_extraction
from core.verifier import run_verification
from core.gap_analyzer import (
    load_output, get_gap_rows, get_component_summary,
    get_region_summary, get_year_summary, get_component_rows
)
from core.paths import get_exe_dir, ensure_output_dir

# ---------------------------------------------------------------------------
# Brand palette
# ---------------------------------------------------------------------------
NORDEX_GREEN   = "#0057A8"
NORDEX_GREEN_D = "#003F7D"   # darker shade for hover / active
BG_MAIN        = "#F0F2F5"
BG_CARD        = "#FFFFFF"
BG_HEADER      = "#003F7D"
TEXT_DARK      = "#1D1D1B"
TEXT_LIGHT     = "#FFFFFF"
TEXT_MID       = "#555555"
BORDER         = "#D5D9E0"
SUCCESS        = "#27AE60"
WARNING        = "#E67E22"
ERROR_CLR      = "#E74C3C"

FONT_TITLE  = ("Segoe UI", 18, "bold")
FONT_HEAD   = ("Segoe UI", 10, "bold")
FONT_BODY   = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 8)
FONT_MONO   = ("Consolas", 8)
FONT_BTN    = ("Segoe UI", 10, "bold")
FONT_BTN_SM = ("Segoe UI", 9)


# ---------------------------------------------------------------------------
# Theme helper
# ---------------------------------------------------------------------------

def apply_style():
    """Configure ttk theme for corporate Nordex look."""
    s = ttk.Style()
    try:
        s.theme_use("clam")
    except Exception:
        pass

    # Base
    s.configure(".", background=BG_MAIN, foreground=TEXT_DARK, font=FONT_BODY)

    # Card frame (white rounded-feel)
    s.configure("Card.TFrame", background=BG_CARD, relief="flat")
    s.configure("Card.TLabelframe", background=BG_CARD)
    s.configure("Card.TLabelframe.Label",
                background=BG_CARD, foreground=TEXT_DARK, font=FONT_HEAD)

    # Section header label
    s.configure("Step.TLabel",
                background=BG_CARD, foreground=NORDEX_GREEN, font=FONT_HEAD)
    s.configure("StepNum.TLabel",
                background=NORDEX_GREEN, foreground=TEXT_LIGHT,
                font=("Segoe UI", 9, "bold"), padding=(4, 2))
    s.configure("Body.TLabel", background=BG_CARD, foreground=TEXT_DARK, font=FONT_BODY)
    s.configure("Small.TLabel", background=BG_CARD, foreground=TEXT_MID, font=FONT_SMALL)
    s.configure("Status.TLabel", background=BG_MAIN, foreground=TEXT_MID, font=FONT_SMALL)

    # Main CTA button - green
    s.configure("Extract.TButton",
                background=NORDEX_GREEN, foreground=TEXT_LIGHT,
                font=FONT_BTN, padding=(10, 12),
                relief="flat", borderwidth=0)
    s.map("Extract.TButton",
          background=[("active", NORDEX_GREEN_D), ("disabled", "#AAAAAA")],
          foreground=[("disabled", "#DDDDDD")])

    # Secondary buttons
    s.configure("Action.TButton",
                background=BG_CARD, foreground=NORDEX_GREEN,
                font=FONT_BTN_SM, padding=(8, 8),
                relief="flat", borderwidth=1)
    s.map("Action.TButton",
          background=[("active", "#E3F0FF"), ("disabled", BG_MAIN)],
          foreground=[("disabled", "#AAAAAA")])

    # Progress bar - blue
    s.configure("Green.Horizontal.TProgressbar",
                troughcolor=BORDER, background=NORDEX_GREEN,
                lightcolor=NORDEX_GREEN, darkcolor=NORDEX_GREEN_D,
                bordercolor=BORDER)

    # Entry
    s.configure("TEntry", fieldbackground=BG_CARD, foreground=TEXT_DARK,
                bordercolor=BORDER, relief="flat", padding=6)


# ---------------------------------------------------------------------------
# Year multi-select dropdown widget
# ---------------------------------------------------------------------------

class YearDropdown(tk.Frame):
    """
    Custom multi-select dropdown showing selected years as summary.
    Click the button to expand/collapse a checklist panel.
    """

    AVAILABLE_YEARS = [2026, 2027, 2028, 2029]

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=BG_CARD, **kwargs)

        # Selection state: year -> BooleanVar
        self.vars = {
            yr: tk.BooleanVar(value=(yr != 2029))
            for yr in self.AVAILABLE_YEARS
        }
        self._panel_open = False

        # Toggle button
        self.btn_var = tk.StringVar()
        self._update_btn_text()

        self.toggle_btn = tk.Button(
            self,
            textvariable=self.btn_var,
            anchor="w",
            bg=BG_CARD,
            fg=TEXT_DARK,
            font=FONT_BODY,
            relief="groove",
            bd=1,
            cursor="hand2",
            command=self._toggle_panel,
            padx=8, pady=5,
        )
        self.toggle_btn.pack(fill=tk.X)

        # Dropdown panel (hidden by default)
        self.panel = tk.Frame(
            self, bg=BG_CARD, relief="groove", bd=1
        )

        for yr in self.AVAILABLE_YEARS:
            row = tk.Frame(self.panel, bg=BG_CARD)
            row.pack(fill=tk.X, padx=8, pady=2)

            cb = tk.Checkbutton(
                row,
                text=str(yr),
                variable=self.vars[yr],
                bg=BG_CARD,
                fg=TEXT_DARK,
                activebackground=BG_CARD,
                selectcolor=NORDEX_GREEN,
                font=FONT_BODY,
                command=self._update_btn_text,
                cursor="hand2",
            )
            cb.pack(side=tk.LEFT)

    def _toggle_panel(self):
        if self._panel_open:
            self.panel.pack_forget()
        else:
            self.panel.pack(fill=tk.X)
        self._panel_open = not self._panel_open
        self._update_btn_text()

    def _update_btn_text(self):
        selected = [str(yr) for yr, v in self.vars.items() if v.get()]
        arrow = "▲" if self._panel_open else "▼"
        if selected:
            self.btn_var.set(f"  {', '.join(selected)}   {arrow}")
        else:
            self.btn_var.set(f"  (Ninguno seleccionado)   {arrow}")

    def get_selected(self) -> list[int]:
        return [yr for yr, v in self.vars.items() if v.get()]


# ---------------------------------------------------------------------------
# Gap Analyzer Window
# ---------------------------------------------------------------------------

class GapAnalyzerWindow(tk.Toplevel):
    """Separate window for gap analysis."""

    def __init__(self, parent, output_path: Path):
        super().__init__(parent)
        self.title("Analisis de Gaps  —  Filas sin costo")
        self.geometry("300x350")
        self.configure(bg=BG_MAIN)
        self.output_path = Path(output_path)
        self.df = None
        self.gap_df = None
        self.components_list = []

        self._build_loading_state()
        self.update()

        threading.Thread(target=self._load_data_bg, daemon=True).start()

    def _build_loading_state(self):
        frame = tk.Frame(self, bg=BG_MAIN)
        frame.pack(expand=True)
        tk.Label(frame, text="Cargando datos...",
                 bg=BG_MAIN, fg=TEXT_MID, font=FONT_BODY).pack(pady=10)
        self.spinner = ttk.Progressbar(frame, mode='indeterminate', length=200)
        self.spinner.pack(pady=5)
        self.spinner.start(15)

    def _load_data_bg(self):
        try:
            self.df = load_output(self.output_path)
            self.gap_df = get_gap_rows(self.df)
            self.after(0, self._build_layout)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Error al cargar: {e}"))
            self.after(0, self.destroy)

    def _build_layout(self):
        for child in self.winfo_children():
            child.destroy()

        # Header
        hdr = tk.Frame(self, bg=BG_HEADER, height=46)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="Analisis de Gaps",
                 bg=BG_HEADER, fg=TEXT_LIGHT,
                 font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=16)
        gap_count = len(self.gap_df)
        total = len(self.df)
        tk.Label(hdr, text=f"{gap_count:,} filas sin costo de {total:,} totales",
                 bg=BG_HEADER, fg="#AAAAAA", font=FONT_SMALL).pack(side=tk.LEFT, padx=4)

        # Main container
        container = tk.Frame(self, bg=BG_MAIN)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ---- Left panel: Components ----
        left_card = tk.Frame(container, bg=BG_CARD, bd=0,
                             highlightbackground=BORDER, highlightthickness=1)
        left_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=False,
                       padx=(0, 8), ipadx=2)
        left_card.pack_propagate(False)
        left_card.config(width=270)

        tk.Label(left_card, text="  COMPONENTES",
                 bg=NORDEX_GREEN, fg=TEXT_LIGHT,
                 font=FONT_HEAD, anchor="w", height=2).pack(fill=tk.X)

        hint = tk.Label(left_card,
                        text="Selecciona para ver detalles",
                        bg=BG_CARD, fg=TEXT_MID, font=FONT_SMALL)
        hint.pack(anchor="w", padx=8, pady=(4, 2))

        list_frame = tk.Frame(left_card, bg=BG_CARD)
        list_frame.pack(fill=tk.BOTH, expand=True)

        sb_l = ttk.Scrollbar(list_frame)
        sb_l.pack(side=tk.RIGHT, fill=tk.Y)

        self.components_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=sb_l.set,
            selectmode=tk.SINGLE,
            bg=BG_CARD,
            fg=TEXT_DARK,
            selectbackground=NORDEX_GREEN,
            selectforeground=TEXT_LIGHT,
            font=FONT_BODY,
            relief="flat",
            bd=0,
            activestyle="none",
            cursor="hand2",
        )
        self.components_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.components_listbox.bind("<<ListboxSelect>>", self._on_component_select)
        sb_l.config(command=self.components_listbox.yview)

        comp_summary = get_component_summary(self.gap_df)
        for comp, count in comp_summary:
            self.components_listbox.insert(tk.END, f"  {comp}  ({count:,})")
            self.components_list.append(comp)

        # ---- Right panel: Detail table ----
        right_card = tk.Frame(container, bg=BG_CARD, bd=0,
                              highlightbackground=BORDER, highlightthickness=1)
        right_card.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tk.Label(right_card, text="  DETALLE",
                 bg=NORDEX_GREEN, fg=TEXT_LIGHT,
                 font=FONT_HEAD, anchor="w", height=2).pack(fill=tk.X)

        self.detail_label = tk.Label(
            right_card,
            text="Selecciona un componente en la lista",
            bg=BG_CARD, fg=TEXT_MID, font=FONT_SMALL, anchor="w"
        )
        self.detail_label.pack(fill=tk.X, padx=8, pady=(4, 2))

        tree_frame = tk.Frame(right_card, bg=BG_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        sb_ty = ttk.Scrollbar(tree_frame)
        sb_ty.pack(side=tk.RIGHT, fill=tk.Y)
        sb_tx = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        sb_tx.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=["Key", "Brand", "Year", "Region", "C1", "C2"],
            height=22,
            yscrollcommand=sb_ty.set,
            xscrollcommand=sb_tx.set,
            show="headings",
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_ty.config(command=self.tree.yview)
        sb_tx.config(command=self.tree.xview)

        self.tree.column("Key",    anchor=tk.W,      width=280)
        self.tree.column("Brand",  anchor=tk.CENTER, width=60)
        self.tree.column("Year",   anchor=tk.CENTER, width=60)
        self.tree.column("Region", anchor=tk.CENTER, width=90)
        self.tree.column("C1",     anchor=tk.E,      width=100)
        self.tree.column("C2",     anchor=tk.E,      width=100)

        self.tree.heading("Key",    text="Torre")
        self.tree.heading("Brand",  text="Marca")
        self.tree.heading("Year",   text="Año")
        self.tree.heading("Region", text="Region")
        self.tree.heading("C1",     text="Cost_Currency1")
        self.tree.heading("C2",     text="Cost_Currency2")

        # Zebra stripes
        self.tree.tag_configure("odd",  background="#F9FFF9")
        self.tree.tag_configure("even", background=BG_CARD)

        # ---- Summary strip ----
        summary_frame = tk.Frame(self, bg=BG_MAIN)
        summary_frame.pack(fill=tk.X, padx=10, pady=(0, 8))

        reg_summary = get_region_summary(self.gap_df)
        yr_summary  = get_year_summary(self.gap_df)

        reg_txt = "  ".join([f"{r}: {c:,}" for r, c in reg_summary[:6]])
        yr_txt  = "  ".join([f"{int(y)}: {c:,}" for y, c in yr_summary])

        tk.Label(summary_frame,
                 text=f"Por region — {reg_txt}",
                 bg=BG_MAIN, fg=TEXT_MID, font=FONT_SMALL).pack(anchor="w")
        tk.Label(summary_frame,
                 text=f"Por año     — {yr_txt}",
                 bg=BG_MAIN, fg=TEXT_MID, font=FONT_SMALL).pack(anchor="w")

        ttk.Button(summary_frame, text="Cerrar",
                   command=self.destroy, style="Action.TButton").pack(
                       anchor="e", pady=4)

    def _on_component_select(self, event):
        if not self.components_listbox.curselection():
            return
        idx = self.components_listbox.curselection()[0]
        component = self.components_list[idx]

        comp_rows = get_component_rows(self.gap_df, component, max_rows=500)
        total_comp = len(self.gap_df[self.gap_df["Component"] == component])

        self.detail_label.config(
            text=f"  {component}  —  {len(comp_rows):,} filas"
                 + (f"  (mostrando {len(comp_rows):,} de {total_comp:,})"
                    if total_comp > len(comp_rows) else "")
        )

        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, (_, row) in enumerate(comp_rows.iterrows()):
            tag = "odd" if i % 2 else "even"
            c1 = f"{row['Cost_Currency1']:,.2f}" if pd.notna(row['Cost_Currency1']) else "—"
            c2 = f"{row['Cost_Currency2']:,.2f}" if pd.notna(row['Cost_Currency2']) else "—"
            self.tree.insert("", tk.END, tags=(tag,), values=(
                row["Key"],
                row["Brand"],
                int(row["Year_Production"]),
                row["Region"],
                c1, c2,
            ))


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------

class App(tk.Tk):
    """Main application window — Nordex corporate style."""

    def __init__(self):
        super().__init__()
        self.title("NDX Sales Calculation Tool  v1.0")
        self.geometry("600x700")
        self.minsize(560, 620)
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)

        apply_style()

        # State
        self.input_path   = tk.StringVar()
        self.output_path  = tk.StringVar()
        self.current_output_path = None
        self.progress_queue = queue.Queue()
        self._status_msg  = tk.StringVar(value="Listo")

        self._setup_layout()

    # ------------------------------------------------------------------ layout

    def _setup_layout(self):
        # ---- Header banner ----
        header = tk.Frame(self, bg=BG_HEADER, height=64)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(
            header,
            text="NORDEX",
            bg=BG_HEADER, fg=NORDEX_GREEN,
            font=("Segoe UI", 16, "bold"),
            padx=16,
        ).pack(side=tk.LEFT)

        tk.Label(
            header,
            text="Sales Calculation Extractor",
            bg=BG_HEADER, fg="#CCCCCC",
            font=("Segoe UI", 11),
        ).pack(side=tk.LEFT)

        tk.Label(
            header,
            text="v1.0",
            bg=BG_HEADER, fg="#666666",
            font=FONT_SMALL,
            padx=12,
        ).pack(side=tk.RIGHT, anchor="s", pady=6)

        # ---- Scrollable body ----
        body = tk.Frame(self, bg=BG_MAIN)
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)

        # ---- Step 1: Input file ----
        self._make_step(body, "1", "ARCHIVO FUENTE")
        card1 = tk.Frame(body, bg=BG_CARD,
                         highlightbackground=BORDER, highlightthickness=1)
        card1.pack(fill=tk.X, pady=(0, 10))

        row1 = tk.Frame(card1, bg=BG_CARD)
        row1.pack(fill=tk.X, padx=10, pady=10)

        self.input_entry = tk.Entry(
            row1, textvariable=self.input_path,
            font=FONT_BODY, bg=BG_CARD, fg=TEXT_DARK,
            relief="groove", bd=1,
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)

        tk.Button(
            row1, text="Seleccionar...",
            bg=NORDEX_GREEN, fg=TEXT_LIGHT,
            font=FONT_BTN_SM, relief="flat",
            activebackground=NORDEX_GREEN_D,
            activeforeground=TEXT_LIGHT,
            cursor="hand2", padx=10, pady=5,
            command=self._browse_input,
        ).pack(side=tk.LEFT, padx=(8, 0))

        self.file_hint = tk.Label(
            card1,
            text="Selecciona el archivo Excel de entrada (.xlsx)",
            bg=BG_CARD, fg=TEXT_MID, font=FONT_SMALL,
        )
        self.file_hint.pack(anchor="w", padx=10, pady=(0, 8))

        # ---- Step 2: Years ----
        self._make_step(body, "2", "ANOS A PROCESAR")
        card2 = tk.Frame(body, bg=BG_CARD,
                         highlightbackground=BORDER, highlightthickness=1)
        card2.pack(fill=tk.X, pady=(0, 10))

        year_inner = tk.Frame(card2, bg=BG_CARD)
        year_inner.pack(fill=tk.X, padx=10, pady=10)

        self.year_dropdown = YearDropdown(year_inner)
        self.year_dropdown.pack(fill=tk.X)

        tk.Label(
            card2,
            text="Los anos no disponibles en el archivo seran ignorados",
            bg=BG_CARD, fg=TEXT_MID, font=FONT_SMALL,
        ).pack(anchor="w", padx=10, pady=(0, 8))

        # ---- Step 3: Output file ----
        self._make_step(body, "3", "GUARDAR RESULTADO EN")
        card3 = tk.Frame(body, bg=BG_CARD,
                         highlightbackground=BORDER, highlightthickness=1)
        card3.pack(fill=tk.X, pady=(0, 10))

        row3 = tk.Frame(card3, bg=BG_CARD)
        row3.pack(fill=tk.X, padx=10, pady=10)

        tk.Entry(
            row3, textvariable=self.output_path,
            font=FONT_BODY, bg=BG_CARD, fg=TEXT_DARK,
            relief="groove", bd=1,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)

        tk.Button(
            row3, text="Guardar en...",
            bg=BG_CARD, fg=NORDEX_GREEN,
            font=FONT_BTN_SM, relief="groove",
            activebackground="#E3F0FF",
            cursor="hand2", padx=10, pady=5,
            command=self._browse_output,
        ).pack(side=tk.LEFT, padx=(8, 0))

        # ---- Extract button ----
        btn_frame = tk.Frame(body, bg=BG_MAIN)
        btn_frame.pack(fill=tk.X, pady=(4, 8))

        self.extract_btn = tk.Button(
            btn_frame,
            text="  EXTRAER DATOS",
            bg=NORDEX_GREEN, fg=TEXT_LIGHT,
            font=("Segoe UI", 12, "bold"),
            relief="flat",
            activebackground=NORDEX_GREEN_D,
            activeforeground=TEXT_LIGHT,
            cursor="hand2",
            pady=12,
            command=self._run_extraction,
        )
        self.extract_btn.pack(fill=tk.X)

        # ---- Progress ----
        prog_card = tk.Frame(body, bg=BG_CARD,
                             highlightbackground=BORDER, highlightthickness=1)
        prog_card.pack(fill=tk.X, pady=(0, 10))

        prog_row = tk.Frame(prog_card, bg=BG_CARD)
        prog_row.pack(fill=tk.X, padx=10, pady=(10, 4))

        tk.Label(prog_row, text="Progreso",
                 bg=BG_CARD, fg=TEXT_DARK, font=FONT_HEAD).pack(side=tk.LEFT)

        self.pct_label = tk.Label(prog_row, text="0%",
                                  bg=BG_CARD, fg=NORDEX_GREEN,
                                  font=("Segoe UI", 9, "bold"))
        self.pct_label.pack(side=tk.RIGHT)

        self.progress = ttk.Progressbar(
            prog_card, mode='determinate', maximum=100,
            style="Green.Horizontal.TProgressbar",
        )
        self.progress.pack(fill=tk.X, padx=10, pady=(0, 8))

        self.status_label = tk.Label(
            prog_card, textvariable=self._status_msg,
            bg=BG_CARD, fg=TEXT_MID, font=FONT_SMALL,
        )
        self.status_label.pack(anchor="w", padx=10, pady=(0, 8))

        # ---- Post-extraction actions ----
        actions_frame = tk.Frame(body, bg=BG_MAIN)
        actions_frame.pack(fill=tk.X, pady=(0, 8))

        self.verify_btn = tk.Button(
            actions_frame,
            text="Verificar Output",
            bg=BG_CARD, fg=TEXT_DARK,
            font=FONT_BTN_SM, relief="groove",
            activebackground="#E3F0FF",
            cursor="hand2", pady=8,
            state=tk.DISABLED,
            command=self._run_verification,
        )
        self.verify_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.gaps_btn = tk.Button(
            actions_frame,
            text="Analizar Gaps",
            bg=BG_CARD, fg=TEXT_DARK,
            font=FONT_BTN_SM, relief="groove",
            activebackground="#E3F0FF",
            cursor="hand2", pady=8,
            state=tk.DISABLED,
            command=self._open_gap_analyzer,
        )
        self.gaps_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ---- Log ----
        log_header = tk.Frame(body, bg=BG_MAIN)
        log_header.pack(fill=tk.X)
        tk.Label(log_header, text="REGISTRO DE ACTIVIDAD",
                 bg=BG_MAIN, fg=TEXT_MID, font=FONT_HEAD).pack(anchor="w")

        log_card = tk.Frame(body, bg=BG_CARD,
                            highlightbackground=BORDER, highlightthickness=1)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.log = scrolledtext.ScrolledText(
            log_card, height=10,
            state=tk.DISABLED, wrap=tk.WORD,
            bg="#0D1117", fg="#C9D1D9",
            font=FONT_MONO,
            insertbackground="#C9D1D9",
            relief="flat", bd=0,
            padx=8, pady=8,
        )
        self.log.pack(fill=tk.BOTH, expand=True)

        # Color tags for log
        self.log.tag_configure("info",    foreground="#58A6FF")
        self.log.tag_configure("success", foreground="#3FB950")
        self.log.tag_configure("warning", foreground="#D29922")
        self.log.tag_configure("error",   foreground="#F85149")

    def _make_step(self, parent, number: str, label: str):
        """Render a step header row."""
        row = tk.Frame(parent, bg=BG_MAIN)
        row.pack(fill=tk.X, pady=(4, 2))

        tk.Label(
            row, text=f" {number} ",
            bg=NORDEX_GREEN, fg=TEXT_LIGHT,
            font=("Segoe UI", 8, "bold"),
        ).pack(side=tk.LEFT)

        tk.Label(
            row, text=f"  {label}",
            bg=BG_MAIN, fg=TEXT_DARK,
            font=FONT_HEAD,
        ).pack(side=tk.LEFT)

    # ----------------------------------------------------------------- actions

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo de entrada",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)
            self.file_hint.config(
                text=f"  {Path(path).name}", fg=NORDEX_GREEN
            )
            # Suggest output path
            base_path = Path(path).parent
            date_str  = datetime.now().strftime("%d%m%y")
            suggested = base_path / f"SalesCalc_{date_str}.xlsx"
            self.output_path.set(str(suggested))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Guardar resultado como",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)

    def _get_selected_years(self) -> list[int]:
        return self.year_dropdown.get_selected()

    def _validate_inputs(self) -> bool:
        if not self.input_path.get():
            messagebox.showerror("Error de validacion",
                                 "Selecciona el archivo de entrada.")
            return False
        if not Path(self.input_path.get()).exists():
            messagebox.showerror("Error de validacion",
                                 "El archivo de entrada no existe.")
            return False
        if not self._get_selected_years():
            messagebox.showerror("Error de validacion",
                                 "Selecciona al menos un ano.")
            return False
        if not self.output_path.get():
            messagebox.showerror("Error de validacion",
                                 "Especifica la ruta de salida.")
            return False
        return True

    # ---------------------------------------------------------------- threads

    def _run_extraction(self):
        if not self._validate_inputs():
            return

        self.extract_btn.config(state=tk.DISABLED, bg="#AAAAAA")
        self.verify_btn.config(state=tk.DISABLED)
        self.gaps_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.pct_label.config(text="0%")
        self._status_msg.set("Procesando...")
        self._log_append(">>> Extraccion iniciada", tag="info")
        self._log_append(
            f"    Archivo: {Path(self.input_path.get()).name}", tag="info"
        )
        self._log_append(
            f"    Anos: {self._get_selected_years()}", tag="info"
        )

        threading.Thread(target=self._extraction_worker, daemon=True).start()
        self._poll_queue()

    def _extraction_worker(self):
        try:
            run_extraction(
                input_path=self.input_path.get(),
                output_path=self.output_path.get(),
                target_years=self._get_selected_years(),
                progress_callback=self._enqueue_progress,
                log_callback=self._enqueue_log,
            )
            self.progress_queue.put(("DONE", self.output_path.get()))
        except Exception as e:
            self.progress_queue.put(("ERROR", str(e)))

    def _enqueue_progress(self, pct: int, msg: str):
        self.progress_queue.put(("PROGRESS", pct, msg))

    def _enqueue_log(self, msg: str):
        self.progress_queue.put(("LOG", msg))

    def _poll_queue(self):
        try:
            while True:
                msg = self.progress_queue.get_nowait()

                if msg[0] == "PROGRESS":
                    _, pct, text = msg
                    self.progress['value'] = pct
                    self.pct_label.config(text=f"{pct}%")
                    self._log_append(f"[{pct:3d}%] {text}", tag="info")

                elif msg[0] == "LOG":
                    _, text = msg
                    self._log_append(text)

                elif msg[0] == "DONE":
                    _, output_path = msg
                    self.current_output_path = output_path
                    self.progress['value'] = 100
                    self.pct_label.config(text="100%", fg=SUCCESS)
                    self._status_msg.set(
                        f"Completado  —  {Path(output_path).name}")
                    self.extract_btn.config(state=tk.NORMAL, bg=NORDEX_GREEN)
                    self.verify_btn.config(state=tk.NORMAL)
                    self.gaps_btn.config(state=tk.NORMAL)
                    self._log_append(
                        f">>> Archivo generado: {output_path}", tag="success")
                    return

                elif msg[0] == "ERROR":
                    _, error_msg = msg
                    self._status_msg.set("Error en extraccion")
                    self.extract_btn.config(state=tk.NORMAL, bg=NORDEX_GREEN)
                    self._log_append(f"ERROR: {error_msg}", tag="error")
                    messagebox.showerror("Error en extraccion", error_msg)
                    return

        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    def _log_append(self, text: str, tag: str = ""):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {text}\n"
        self.log.config(state=tk.NORMAL)
        if tag:
            self.log.insert(tk.END, line, tag)
        else:
            self.log.insert(tk.END, line)
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _run_verification(self):
        if not self.current_output_path:
            messagebox.showerror("Error", "No hay output para verificar.")
            return

        self.verify_btn.config(state=tk.DISABLED)
        self._log_append(">>> Verificacion iniciada", tag="info")
        self._status_msg.set("Verificando...")

        def verify_worker():
            try:
                run_verification(
                    output_path=self.current_output_path,
                    selected_years=self._get_selected_years(),
                    canonical_regions=[
                        "Europe", "Germany", "Turkey", "Turkey_DOM",
                        "Asia", "China", "Poland", "Italy", "Greece", "CAN", "US"
                    ],
                    log_callback=self._enqueue_log,
                )
                self.progress_queue.put(("VERIFY_DONE", None))
            except Exception as e:
                self.progress_queue.put(("VERIFY_ERROR", str(e)))

        threading.Thread(target=verify_worker, daemon=True).start()
        self._poll_queue_verify()

    def _poll_queue_verify(self):
        try:
            while True:
                msg = self.progress_queue.get_nowait()
                if msg[0] == "LOG":
                    self._log_append(msg[1])
                elif msg[0] == "VERIFY_DONE":
                    self.verify_btn.config(state=tk.NORMAL)
                    self._status_msg.set("Verificacion completada")
                    return
                elif msg[0] == "VERIFY_ERROR":
                    self._log_append(f"ERROR: {msg[1]}", tag="error")
                    messagebox.showerror("Error en verificacion", msg[1])
                    self.verify_btn.config(state=tk.NORMAL)
                    self._status_msg.set("Error en verificacion")
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_queue_verify)

    def _open_gap_analyzer(self):
        if not self.current_output_path:
            messagebox.showerror("Error", "No hay output para analizar.")
            return
        GapAnalyzerWindow(self, self.current_output_path)


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
