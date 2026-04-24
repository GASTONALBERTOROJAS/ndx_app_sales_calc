"""
NDX Sales Calculation Tool - GUI Application (tkinter).

Main entry point for the graphical user interface.
Integrates core extraction, verification, and gap analysis.
"""

import sys
import queue
import threading
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


class GapAnalyzerWindow(tk.Toplevel):
    """Separate window for gap analysis (component selection + data viewing)."""

    def __init__(self, parent, output_path: Path):
        super().__init__(parent)
        self.title("Analisis de Gaps — Filas sin costo")
        self.geometry("1000x600")
        self.output_path = Path(output_path)
        self.df = None
        self.gap_df = None
        self.components_list = []

        # Loading state
        self._build_loading_state()
        self.update()

        # Load data in separate thread
        threading.Thread(target=self._load_data_bg, daemon=True).start()

    def _build_loading_state(self):
        """Show loading indicator."""
        frame = ttk.Frame(self)
        frame.pack(expand=True)
        ttk.Label(frame, text="Cargando datos...").pack(pady=10)
        self.progress = ttk.Progressbar(frame, mode='indeterminate')
        self.progress.pack(pady=5)
        self.progress.start()

    def _load_data_bg(self):
        """Load DataFrame in background."""
        try:
            self.df = load_output(self.output_path)
            self.gap_df = get_gap_rows(self.df)
            self.after(0, self._build_layout)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Error al cargar: {e}"))
            self.destroy()

    def _build_layout(self):
        """Build main UI after data loaded."""
        # Clear loading indicator
        for child in self.winfo_children():
            child.destroy()

        # Main container
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left panel: Component list
        left_frame = ttk.Frame(container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5))

        ttk.Label(left_frame, text=f"Componentes\n({len(self.gap_df)} filas sin costo)").pack()

        # Listbox with scrollbar
        scrollbar = ttk.Scrollbar(left_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.components_listbox = tk.Listbox(
            left_frame, yscrollcommand=scrollbar.set, width=40
        )
        self.components_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.components_listbox.bind("<<ListboxSelect>>", self._on_component_select)
        scrollbar.config(command=self.components_listbox.yview)

        # Populate listbox
        comp_summary = get_component_summary(self.gap_df)
        for idx, (comp, count) in enumerate(comp_summary):
            self.components_listbox.insert(tk.END, f"{comp} ({count})")
            self.components_list.append(comp)

        # Right panel: Treeview with data
        right_frame = ttk.Frame(container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(right_frame, text="Detalles del componente").pack()

        # Treeview with scrollbars
        tree_frame = ttk.Frame(right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        tree_scroll_y = ttk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=["Component", "Key", "Brand", "Year", "Region", "C1", "C2"],
            height=20,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        # Configure columns
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("Component", anchor=tk.W, width=200)
        self.tree.column("Key", anchor=tk.W, width=150)
        self.tree.column("Brand", anchor=tk.CENTER, width=50)
        self.tree.column("Year", anchor=tk.CENTER, width=50)
        self.tree.column("Region", anchor=tk.CENTER, width=80)
        self.tree.column("C1", anchor=tk.RIGHT, width=80)
        self.tree.column("C2", anchor=tk.RIGHT, width=80)

        self.tree.heading("#0", text="")
        self.tree.heading("Component", text="Componente")
        self.tree.heading("Key", text="Torre")
        self.tree.heading("Brand", text="Marca")
        self.tree.heading("Year", text="Año")
        self.tree.heading("Region", text="Region")
        self.tree.heading("C1", text="Cost_C1")
        self.tree.heading("C2", text="Cost_C2")

        # Bottom: Summary info
        bottom_frame = ttk.LabelFrame(self, text="Distribución de Gaps")
        bottom_frame.pack(fill=tk.X, padx=5, pady=5)

        reg_summary = get_region_summary(self.gap_df)
        yr_summary = get_year_summary(self.gap_df)

        reg_text = ", ".join([f"{r}:{c}" for r, c in reg_summary[:5]])
        yr_text = ", ".join([f"{int(y)}:{c}" for y, c in yr_summary])

        ttk.Label(bottom_frame, text=f"Regiones: {reg_text}").pack(anchor=tk.W, padx=5)
        ttk.Label(bottom_frame, text=f"Años: {yr_text}").pack(anchor=tk.W, padx=5)

        # Close button
        ttk.Button(bottom_frame, text="Cerrar", command=self.destroy).pack(anchor=tk.E, padx=5, pady=5)

    def _on_component_select(self, event):
        """Populate Treeview when component selected."""
        if not self.components_listbox.curselection():
            return
        idx = self.components_listbox.curselection()[0]
        component = self.components_list[idx]

        # Get rows for this component
        comp_rows = get_component_rows(self.gap_df, component, max_rows=500)

        # Clear and populate Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in comp_rows.iterrows():
            values = (
                row["Component"],
                row["Key"],
                row["Brand"],
                int(row["Year_Production"]),
                row["Region"],
                f"{row['Cost_Currency1']:.2f}" if pd.notna(row['Cost_Currency1']) else "",
                f"{row['Cost_Currency2']:.2f}" if pd.notna(row['Cost_Currency2']) else "",
            )
            self.tree.insert("", tk.END, values=values)


class App(tk.Tk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("NDX Sales Calculation Tool  v1.0")
        self.geometry("650x850")
        self.resizable(False, True)

        # State
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.current_output_path = None
        self.progress_queue = queue.Queue()

        # Year checkboxes (BooleanVar)
        self.year_2026 = tk.BooleanVar(value=True)
        self.year_2027 = tk.BooleanVar(value=True)
        self.year_2028 = tk.BooleanVar(value=True)
        self.year_2029 = tk.BooleanVar(value=False)

        self._setup_layout()

    def _setup_layout(self):
        """Build all UI elements."""
        # --- Input file frame ---
        input_frame = ttk.LabelFrame(self, text="[PASO 1] Archivo de entrada", padding=10)
        input_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Entry(input_frame, textvariable=self.input_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="Buscar...", command=self._browse_input).pack(side=tk.LEFT, padx=5)

        # --- Years frame ---
        years_frame = ttk.LabelFrame(self, text="[PASO 2] Años a procesar", padding=10)
        years_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Checkbutton(years_frame, text="2026", variable=self.year_2026).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(years_frame, text="2027", variable=self.year_2027).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(years_frame, text="2028", variable=self.year_2028).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(years_frame, text="2029", variable=self.year_2029).pack(side=tk.LEFT, padx=10)

        # --- Output file frame ---
        output_frame = ttk.LabelFrame(self, text="[PASO 3] Guardar resultado en...", padding=10)
        output_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_frame, text="Guardar en", command=self._browse_output).pack(side=tk.LEFT, padx=5)

        # --- Extract button ---
        extract_frame = ttk.Frame(self)
        extract_frame.pack(fill=tk.X, padx=10, pady=10)

        self.extract_btn = ttk.Button(
            extract_frame, text="EXTRAER DATOS", command=self._run_extraction
        )
        self.extract_btn.pack(fill=tk.X, ipady=10)

        # --- Progress bar ---
        self.progress = ttk.Progressbar(self, mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        progress_label_frame = ttk.Frame(self)
        progress_label_frame.pack(fill=tk.X, padx=10)
        self.progress_label = ttk.Label(progress_label_frame, text="0%")
        self.progress_label.pack(anchor=tk.E)

        # --- Post-extraction buttons ---
        actions_frame = ttk.Frame(self)
        actions_frame.pack(fill=tk.X, padx=10, pady=10)

        self.verify_btn = ttk.Button(
            actions_frame, text="Verificar Output", command=self._run_verification, state=tk.DISABLED
        )
        self.verify_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.gaps_btn = ttk.Button(
            actions_frame, text="Analizar Gaps", command=self._open_gap_analyzer, state=tk.DISABLED
        )
        self.gaps_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # --- Log frame ---
        log_frame = ttk.LabelFrame(self, text="Log:", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.log = scrolledtext.ScrolledText(log_frame, height=15, state=tk.DISABLED, wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True)

    def _browse_input(self):
        """File dialog for input Excel."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo de entrada",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)
            # Suggest output path in same folder
            base_path = Path(path).parent
            today = datetime.now()
            date_str = today.strftime("%d%m%y")
            suggested_output = base_path / f"SalesCalc_{date_str}.xlsx"
            self.output_path.set(str(suggested_output))

    def _browse_output(self):
        """File dialog for output Excel."""
        path = filedialog.asksaveasfilename(
            title="Guardar resultado como",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)

    def _get_selected_years(self) -> list[int]:
        """Return list of selected years."""
        years = []
        if self.year_2026.get():
            years.append(2026)
        if self.year_2027.get():
            years.append(2027)
        if self.year_2028.get():
            years.append(2028)
        if self.year_2029.get():
            years.append(2029)
        return years

    def _validate_inputs(self) -> bool:
        """Validate all inputs before extraction."""
        if not self.input_path.get():
            messagebox.showerror("Error", "Seleccione archivo de entrada")
            return False
        if not Path(self.input_path.get()).exists():
            messagebox.showerror("Error", "Archivo de entrada no encontrado")
            return False
        if not self._get_selected_years():
            messagebox.showerror("Error", "Seleccione al menos un año")
            return False
        if not self.output_path.get():
            messagebox.showerror("Error", "Especifique ruta de salida")
            return False
        return True

    def _run_extraction(self):
        """Run extraction in background thread."""
        if not self._validate_inputs():
            return

        # Disable buttons, reset progress
        self.extract_btn.config(state=tk.DISABLED)
        self.verify_btn.config(state=tk.DISABLED)
        self.gaps_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.progress_label.config(text="0%")
        self._log_append("")

        # Start worker thread
        threading.Thread(target=self._extraction_worker, daemon=True).start()
        # Start polling queue
        self._poll_queue()

    def _extraction_worker(self):
        """Run extraction in background."""
        try:
            result = run_extraction(
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
        """Callback for progress updates."""
        self.progress_queue.put(("PROGRESS", pct, msg))

    def _enqueue_log(self, msg: str):
        """Callback for log messages."""
        self.progress_queue.put(("LOG", msg))

    def _poll_queue(self):
        """Poll queue for messages (called recursively with after)."""
        try:
            while True:
                msg = self.progress_queue.get_nowait()

                if msg[0] == "PROGRESS":
                    _, pct, text = msg
                    self.progress['value'] = pct
                    self.progress_label.config(text=f"{pct}%")
                    self._log_append(text)

                elif msg[0] == "LOG":
                    _, text = msg
                    self._log_append(text)

                elif msg[0] == "DONE":
                    _, output_path = msg
                    self.current_output_path = output_path
                    self.progress['value'] = 100
                    self.progress_label.config(text="100%")
                    self.extract_btn.config(state=tk.NORMAL)
                    self.verify_btn.config(state=tk.NORMAL)
                    self.gaps_btn.config(state=tk.NORMAL)
                    return

                elif msg[0] == "ERROR":
                    _, error_msg = msg
                    messagebox.showerror("Error en extracción", error_msg)
                    self.extract_btn.config(state=tk.NORMAL)
                    self.verify_btn.config(state=tk.DISABLED)
                    self.gaps_btn.config(state=tk.DISABLED)
                    return

        except queue.Empty:
            pass

        # Schedule next poll
        self.after(100, self._poll_queue)

    def _log_append(self, text: str):
        """Append text to log widget."""
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _run_verification(self):
        """Run verification in background."""
        if not self.current_output_path:
            messagebox.showerror("Error", "No hay output para verificar")
            return

        self.verify_btn.config(state=tk.DISABLED)
        self._log_append("\n--- Iniciando verificación ---")

        def verify_worker():
            try:
                result = run_verification(
                    output_path=self.current_output_path,
                    selected_years=self._get_selected_years(),
                    canonical_regions=[
                        "Europe", "Germany", "Turkey", "Turkey_DOM",
                        "Asia", "China", "Poland", "Italy", "Greece", "CAN", "US"
                    ],
                    log_callback=self._enqueue_log,
                )
                self.progress_queue.put(("VERIFY_DONE", result))
            except Exception as e:
                self.progress_queue.put(("VERIFY_ERROR", str(e)))

        threading.Thread(target=verify_worker, daemon=True).start()
        self._poll_queue_verify()

    def _poll_queue_verify(self):
        """Poll for verification results."""
        try:
            while True:
                msg = self.progress_queue.get_nowait()

                if msg[0] == "LOG":
                    _, text = msg
                    self._log_append(text)

                elif msg[0] == "VERIFY_DONE":
                    _, result = msg
                    self.verify_btn.config(state=tk.NORMAL)
                    return

                elif msg[0] == "VERIFY_ERROR":
                    _, error_msg = msg
                    messagebox.showerror("Error en verificación", error_msg)
                    self.verify_btn.config(state=tk.NORMAL)
                    return

        except queue.Empty:
            pass

        self.after(100, self._poll_queue_verify)

    def _open_gap_analyzer(self):
        """Open gap analyzer window."""
        if not self.current_output_path:
            messagebox.showerror("Error", "No hay output para analizar")
            return
        GapAnalyzerWindow(self, self.current_output_path)


if __name__ == "__main__":
    import pandas as pd
    app = App()
    app.mainloop()
