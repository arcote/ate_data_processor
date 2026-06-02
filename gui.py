"""
gui.py
------
Tkinter front-end for the ATE data reporting tool. Ties together the three
processing modules:

  extract_ate_data   — walk the data tree, pull the extracted value per run
  visualize_ate_data — render the review plots
  build_ate_excel    — assemble the Excel workbook with native charts

Workflow:
  1. Pick the input data folder. The app auto-scans it and shows the detected
     layout, parameters and conditions.
  2. Tick the parameters / conditions to include and the plots to generate.
  3. Choose the output formats (xlsx / plot images / csv), the output folder
     and a report name.
  4. Run. Extraction, plotting and workbook building happen on a background
     thread so the UI stays responsive; progress and a live log are shown.

Run:  python gui.py
"""

import os

# Force the non-interactive Agg backend BEFORE anything (directly or via the
# imported modules) pulls in matplotlib.pyplot. Plotting runs on a worker
# thread, where an interactive GUI backend would be unsafe.
import matplotlib
matplotlib.use("Agg")

import queue
import tempfile
import threading
import traceback
from pathlib import Path

import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from PIL import Image, ImageTk

import extract_ate_data as extract
import visualize_ate_data as visualize
import build_ate_excel as build


PLOT_LABELS = {
    "boltzmann": "Boltzmann scatter",
    "box":       "Box-and-whisker",
    "trend":     "Run trend",
    "heatmap":   "Heat map",
    "histogram": "Histogram overlay",
}
PLOT_FILE = {
    "boltzmann": "1_boltzmann_Demo.png",
    "box":       "2_boxplot.png",
    "trend":     "3_run_trend.png",
    "heatmap":   "4_heatmap.png",
    "histogram": "5_histogram.png",
}
IMAGE_FORMATS = ["png", "pdf", "svg"]
THEMES = sorted(visualize.THEMES)
BASE_COLS = ["test_folder", "temp_folder", "run_number", "extracted_value"]


def make_demo_dataframe():
    """Synthetic dataset used to render the plot previews — covers all five
    plots cleanly (1 parameter × 3 conditions × 3 iterations × 8 runs)."""
    rng = np.random.default_rng(0)
    rows = []
    centers = {"COLD": 1.0e-9, "ROOM": 2.0e-9, "HOT": 5.0e-9}
    for test in ("1st test", "2nd test", "3rd test"):
        for cond, c in centers.items():
            for run in range(1, 9):
                val = c * (1 + rng.uniform(-0.08, 0.08))
                rows.append({
                    "test_folder": test,
                    "temp_folder": f"Demo_{cond}",
                    "run_number": run,
                    "extracted_value": val,
                })
    return pd.DataFrame(rows)


class ATEReportApp:
    def __init__(self, root):
        self.root = root
        root.title("ATE Data Report Builder")
        root.geometry("1024x760")
        root.minsize(900, 640)

        # Thread / messaging state
        self.msg_queue = queue.Queue()
        self.worker = None
        self.last_preview = None

        # Tk variables
        self.input_dir_var   = tk.StringVar()
        self.output_dir_var  = tk.StringVar()
        self.report_name_var = tk.StringVar(value="ATE_Report")
        self.status_var      = tk.StringVar(value="Select an input data folder to begin.")
        self.progress_var    = tk.DoubleVar(value=0.0)

        self.param_vars      = {}   # parameter name -> BooleanVar
        self.param_unit_vars = {}   # parameter name -> StringVar (Y-axis unit)
        self.enc_vars        = {}   # enclosure tag (EN4…) -> BooleanVar
        self.cond_vars       = {}   # condition name -> BooleanVar
        self._all_params         = []   # every parameter from the last scan
        self._param_check_cache  = {}   # parameter -> bool (preserved across refilter)
        self._param_unit_cache   = {}   # parameter -> unit string (preserved)
        self.plot_vars  = {p: tk.BooleanVar(value=True) for p in PLOT_LABELS}

        # Temperature value per condition (°C) — drives the Boltzmann x-axis.
        self.temp_vars = {
            "COLD": tk.StringVar(value="-40"),
            "ROOM": tk.StringVar(value="25"),
            "HOT":  tk.StringVar(value="85"),
        }

        self.out_xlsx_var = tk.BooleanVar(value=True)
        self.out_png_var  = tk.BooleanVar(value=True)
        self.out_csv_var  = tk.BooleanVar(value=True)

        self.img_format_var = tk.StringVar(value="png")
        self.dpi_var        = tk.IntVar(value=150)
        self.theme_var      = tk.StringVar(value=visualize.DEFAULT_THEME)
        self.theme_var.trace_add("write", lambda *_: self._refresh_previews())

        # Per-plot custom title (blank = derive from the parameter / plot name)
        self.title_vars = {p: tk.StringVar(value="") for p in PLOT_LABELS}

        # Per-plot Boltzmann controls
        self.limit_mode_var = tk.StringVar(value="sigma")
        self.sigma_k_var    = tk.DoubleVar(value=3.0)
        self.limit_low_var  = tk.StringVar(value="")
        self.limit_high_var = tk.StringVar(value="")
        self.y_unit_var     = tk.StringVar(value="")   # blank = infer
        self.show_table_var = tk.BooleanVar(value=True)
        for v in (self.limit_mode_var, self.sigma_k_var, self.show_table_var):
            v.trace_add("write", lambda *_: self._on_bz_opt_change())

        # Preview state — holds the rendered PhotoImage refs so Tk doesn't
        # garbage-collect them, the cached demo DataFrame, and a temp dir.
        self._preview_dir = Path(tempfile.mkdtemp(prefix="ate_preview_"))
        self._preview_df = make_demo_dataframe()
        self._preview_images = {}      # plot_key -> ImageTk.PhotoImage (shown)
        self._preview_src    = {}      # plot_key -> PIL.Image (full-res source)
        self._preview_labels = {}      # plot_key -> ttk.Label
        self._preview_size   = {}      # plot_key -> last rendered width
        self._preview_thread = None
        self._popouts        = []      # live Toplevel pop-out windows

        self._build_ui()
        self.root.after(100, self._drain_queue)
        self.root.after(50, self._refresh_previews)

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        outer = ttk.Frame(root, padding=10)
        outer.grid(row=0, column=0, sticky="nsew")
        # Left column = 1/3 (everything except plots); right column = 2/3 (plots).
        outer.columnconfigure(0, weight=1, uniform="cols")
        outer.columnconfigure(1, weight=2, uniform="cols")
        outer.rowconfigure(0, weight=1)

        left = ttk.Frame(outer)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.columnconfigure(0, weight=1)
        # The log row absorbs extra vertical space.
        left.rowconfigure(9, weight=1)

        right = ttk.Frame(outer)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self._build_input_row(left, row=0)
        self._build_preview(left, row=1)
        self._build_enclosure_panel(left, row=2)
        self._build_params_panel(left, row=3)
        self._build_conditions_panel(left, row=4)
        self._build_temps_panel(left, row=5)
        self._build_output(left, row=6)
        self._build_run_row(left, row=7)
        self._build_log(left, row=9)

        self._build_plots(right, row=0)

    def _build_input_row(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Input data folder", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        ttk.Entry(frame, textvariable=self.input_dir_var).grid(
            row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(frame, text="Browse…", command=self._choose_input).grid(
            row=0, column=1)
        ttk.Button(frame, text="Re-scan", command=self._start_scan).grid(
            row=0, column=2, padx=(6, 0))

    def _build_preview(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Detected layout", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        self.preview_label = ttk.Label(
            frame, justify="left", anchor="w",
            text="No folder scanned yet.")
        self.preview_label.grid(row=0, column=0, sticky="ew")

    def _build_enclosure_panel(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Enclosures (EN#)", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        self.enc_holder = ttk.Frame(frame)
        self.enc_holder.grid(row=0, column=0, sticky="w")
        self._placeholder(self.enc_holder, "scan a folder to list enclosures")

    def _build_params_panel(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Parameters", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        self.param_holder = ttk.Frame(frame)
        self.param_holder.grid(row=0, column=0, sticky="ew")
        self._placeholder(self.param_holder, "scan a folder to list parameters")

    def _build_conditions_panel(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Conditions", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        self.cond_holder = ttk.Frame(frame)
        self.cond_holder.grid(row=0, column=0, sticky="w")
        self._placeholder(self.cond_holder, "scan a folder to list conditions")

    def _build_temps_panel(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Temperatures (°C)", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        for i, cond in enumerate(("COLD", "ROOM", "HOT")):
            ttk.Label(frame, text=f"{cond}:").grid(row=0, column=i * 2, padx=(0, 4))
            ttk.Entry(frame, textvariable=self.temp_vars[cond], width=7).grid(
                row=0, column=i * 2 + 1, padx=(0, 16))

    def _build_plots(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Plots", padding=8)
        frame.grid(row=row, column=0, sticky="nsew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        self.plots_frame = frame

        # Shared controls: theme, image format, DPI — apply to every plot.
        sub = ttk.Frame(frame)
        sub.grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(sub, text="Theme:").grid(row=0, column=0, padx=(0, 4))
        ttk.Combobox(sub, textvariable=self.theme_var, values=THEMES,
                     width=8, state="readonly").grid(row=0, column=1, padx=(0, 16))
        ttk.Label(sub, text="Image format:").grid(row=0, column=2, padx=(0, 4))
        ttk.Combobox(sub, textvariable=self.img_format_var, values=IMAGE_FORMATS,
                     width=6, state="readonly").grid(row=0, column=3, padx=(0, 16))
        ttk.Label(sub, text="DPI:").grid(row=0, column=4, padx=(0, 4))
        ttk.Spinbox(sub, from_=72, to=600, increment=1, width=6,
                    textvariable=self.dpi_var).grid(row=0, column=5)

        # One tab per plot type. Each tab holds an enable checkbox and a
        # theme-aware preview thumbnail. Future per-plot adjustment controls
        # belong inside these tab frames.
        nb = ttk.Notebook(frame)
        nb.grid(row=1, column=0, sticky="nsew")
        for key, label in PLOT_LABELS.items():
            tab = ttk.Frame(nb, padding=8)
            nb.add(tab, text=label)
            tab.columnconfigure(0, weight=1)

            top = ttk.Frame(tab)
            top.grid(row=0, column=0, sticky="ew")
            top.columnconfigure(0, weight=1)
            ttk.Checkbutton(top, text=f"Include {label}",
                            variable=self.plot_vars[key]).grid(
                row=0, column=0, sticky="w")
            ttk.Button(top, text="Pop out ⤢", width=10,
                       command=lambda k=key: self._popout_preview(k)).grid(
                row=0, column=1, sticky="e")

            title_row = ttk.Frame(tab)
            title_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
            title_row.columnconfigure(1, weight=1)
            ttk.Label(title_row, text="Title:").grid(row=0, column=0, padx=(0, 4))
            title_entry = ttk.Entry(title_row, textvariable=self.title_vars[key])
            title_entry.grid(row=0, column=1, sticky="ew")
            ttk.Label(title_row, text="(blank = parameter name)",
                      foreground="#888888").grid(row=0, column=2, padx=(6, 0))
            for evt in ("<Return>", "<FocusOut>"):
                title_entry.bind(evt, lambda e: self._refresh_previews())

            preview_row = 2
            if key == "boltzmann":
                self._build_boltzmann_controls(tab, row=2)
                preview_row = 3

            preview = ttk.Label(tab, text="rendering preview…",
                                anchor="center", relief="sunken",
                                background="#f0f0f0")
            preview.grid(row=preview_row, column=0, sticky="nsew", pady=(8, 0))
            preview.bind("<Configure>",
                         lambda e, k=key: self._on_preview_configure(k, e))
            preview.bind("<Double-Button-1>",
                         lambda e, k=key: self._popout_preview(k))
            tab.rowconfigure(preview_row, weight=1)
            self._preview_labels[key] = preview

    def _build_boltzmann_controls(self, tab, row):
        ctl = ttk.Frame(tab)
        ctl.grid(row=row, column=0, sticky="ew", pady=(6, 0))

        ttk.Label(ctl, text="Limits:").grid(row=0, column=0, padx=(0, 4))
        self.limit_combo = ttk.Combobox(
            ctl, textvariable=self.limit_mode_var, values=visualize.LIMIT_MODES,
            width=8, state="readonly")
        self.limit_combo.grid(row=0, column=1, padx=(0, 12))

        ttk.Label(ctl, text="k·σ:").grid(row=0, column=2, padx=(0, 4))
        self.sigma_spin = ttk.Spinbox(ctl, from_=0.5, to=10, increment=0.5,
                                      width=5, textvariable=self.sigma_k_var)
        self.sigma_spin.grid(row=0, column=3, padx=(0, 12))

        ttk.Label(ctl, text="Low:").grid(row=0, column=4, padx=(0, 4))
        self.low_entry = ttk.Entry(ctl, textvariable=self.limit_low_var, width=9)
        self.low_entry.grid(row=0, column=5, padx=(0, 8))
        ttk.Label(ctl, text="High:").grid(row=0, column=6, padx=(0, 4))
        self.high_entry = ttk.Entry(ctl, textvariable=self.limit_high_var, width=9)
        self.high_entry.grid(row=0, column=7, padx=(0, 12))

        ttk.Label(ctl, text="Y unit (all):").grid(row=0, column=8, padx=(0, 4))
        self.unit_entry = ttk.Entry(ctl, textvariable=self.y_unit_var, width=6)
        self.unit_entry.grid(row=0, column=9, padx=(0, 4))
        ttk.Label(ctl, text="(per-param above wins)",
                  foreground="#888888").grid(row=0, column=10, padx=(0, 12))

        ttk.Checkbutton(ctl, text="Value table", variable=self.show_table_var).grid(
            row=0, column=11)

        # Re-render the preview when the text fields are committed.
        for entry in (self.low_entry, self.high_entry, self.unit_entry):
            entry.bind("<Return>",  lambda e: self._on_bz_opt_change())
            entry.bind("<FocusOut>", lambda e: self._on_bz_opt_change())

        self._update_bz_control_states()

    def _update_bz_control_states(self):
        """Enable only the limit inputs relevant to the chosen mode."""
        mode = self.limit_mode_var.get()
        self.sigma_spin.configure(state="normal" if mode == "sigma" else "disabled")
        fixed = "normal" if mode == "fixed" else "disabled"
        self.low_entry.configure(state=fixed)
        self.high_entry.configure(state=fixed)

    def _on_bz_opt_change(self):
        self._update_bz_control_states()
        self._refresh_previews()

    def _parse_float(self, s):
        try:
            return float(s)
        except (TypeError, ValueError):
            return None

    def _boltzmann_opts(self):
        """Collect the Boltzmann controls into a kwargs dict for the plotter.

        Per-parameter Y units from the Parameters panel win; the Boltzmann tab's
        Y-unit field acts as a global fallback (stored under the '*' key)."""
        mode = self.limit_mode_var.get()
        title = self.title_vars["boltzmann"].get().strip()
        opts = {
            "limit_mode": mode,
            "sigma_k":    self._parse_float(self.sigma_k_var.get()) or 3.0,
            "show_table": bool(self.show_table_var.get()),
            # Title behavior: blank entry => use the parameter name as the title
            # (per-parameter); non-empty entry => use the entered text verbatim
            # for every parameter.
            "show_title": True,
            "title":      title or None,
        }
        if mode == "fixed":
            opts["limits"] = (self._parse_float(self.limit_low_var.get()),
                              self._parse_float(self.limit_high_var.get()))

        units = {p: v.get().strip() for p, v in self.param_unit_vars.items()
                 if v.get().strip()}
        global_unit = self.y_unit_var.get().strip()
        if global_unit:
            units["*"] = global_unit
        if units:
            opts["units"] = units
        return opts

    # ── Preview rendering ────────────────────────────────────────────────────────

    def _refresh_previews(self):
        """Re-render the per-tab plot previews from the synthetic demo dataset
        using the currently selected theme. Runs on a background thread so the
        UI stays responsive."""
        if self._preview_thread is not None and self._preview_thread.is_alive():
            return
        theme_name = self.theme_var.get()
        bopts = self._boltzmann_opts()
        for lbl in self._preview_labels.values():
            lbl.configure(text=f"rendering ({theme_name}) preview…", image="")
        self._preview_thread = threading.Thread(
            target=self._render_previews_worker, args=(theme_name, bopts),
            daemon=True)
        self._preview_thread.start()

    def _render_previews_worker(self, theme_name, bopts):
        try:
            preview_df = visualize.enrich(self._preview_df)
            # The Boltzmann plot names its file after the parameter; the demo
            # dataset uses a single parameter named "Demo".
            visualize.generate_plots(
                preview_df, list(PLOT_LABELS), str(self._preview_dir),
                fmt="png", dpi=80, theme=theme_name, log=lambda *_: None,
                boltzmann_opts=bopts)
            for key, fname in PLOT_FILE.items():
                path = self._preview_dir / fname
                if path.exists():
                    self._post("preview_image", (key, str(path)))
                else:
                    self._post("preview_missing", key)
        except Exception:
            self._post("error", traceback.format_exc())

    def _build_output(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Output", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(1, weight=1)

        fmts = ttk.Frame(frame)
        fmts.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Label(fmts, text="Formats:").grid(row=0, column=0, padx=(0, 8))
        ttk.Checkbutton(fmts, text="Excel report (.xlsx)",
                        variable=self.out_xlsx_var).grid(row=0, column=1, padx=(0, 12))
        ttk.Checkbutton(fmts, text="Plot images (png/…)",
                        variable=self.out_png_var).grid(row=0, column=2, padx=(0, 12))
        ttk.Checkbutton(fmts, text="Extracted data (.csv)",
                        variable=self.out_csv_var).grid(row=0, column=3)

        ttk.Label(frame, text="Output folder:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_dir_var).grid(
            row=1, column=1, sticky="ew", padx=6)
        ttk.Button(frame, text="Browse…", command=self._choose_output).grid(
            row=1, column=2)

        ttk.Label(frame, text="Report name:").grid(
            row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frame, textvariable=self.report_name_var).grid(
            row=2, column=1, sticky="ew", padx=6, pady=(6, 0))

    def _build_run_row(self, parent, row):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(1, weight=1)

        self.run_button = ttk.Button(frame, text="Run", command=self._start_run)
        self.run_button.grid(row=0, column=0, padx=(0, 8))
        self.progress = ttk.Progressbar(
            frame, maximum=100, variable=self.progress_var)
        self.progress.grid(row=0, column=1, sticky="ew")
        ttk.Label(frame, textvariable=self.status_var).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

    def _build_log(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Log", padding=4)
        frame.grid(row=row, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(frame, command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

    @staticmethod
    def _placeholder(holder, text):
        ttk.Label(holder, text=text, foreground="#888888").grid(
            row=0, column=0, sticky="w")

    # ── Folder pickers ─────────────────────────────────────────────────────────

    def _choose_input(self):
        path = filedialog.askdirectory(title="Select input data folder")
        if path:
            self.input_dir_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(path)
            self._start_scan()

    def _choose_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir_var.set(path)

    # ── Scan (background) ────────────────────────────────────────────────────────

    def _start_scan(self):
        root_dir = self.input_dir_var.get().strip()
        if not root_dir:
            messagebox.showwarning("No folder", "Choose an input data folder first.")
            return
        if not os.path.isdir(root_dir):
            messagebox.showerror("Invalid folder", f"Not a directory:\n{root_dir}")
            return
        if self._busy():
            return
        self._set_status("Scanning…")
        self._log(f"Scanning {root_dir} …")
        self.worker = threading.Thread(
            target=self._scan_worker, args=(root_dir,), daemon=True)
        self.worker.start()

    def _scan_worker(self, root_dir):
        try:
            preview = extract.preview_scan(
                root_dir, log=lambda m: self._post("log", m))
            self._post("preview", preview)
        except Exception:
            self._post("error", traceback.format_exc())

    def _apply_preview(self, preview):
        self.last_preview = preview
        layout = {"A": "Layout A (named test folders)",
                  "B": "Layout B (ACS_BASIC at root)"}.get(
            preview["layout"], "no recognizable layout")
        lines = [
            f"Layout: {layout}",
            f"Run files found: {preview['file_count']}",
            f"Test iterations: {', '.join(preview['test_folders']) or '—'}",
            f"Parameters: {', '.join(preview['parameters']) or '—'}",
            f"Conditions: {', '.join(preview['conditions']) or '—'}",
        ]
        self.preview_label.configure(text="\n".join(lines))

        self._all_params = list(preview["parameters"])
        self._rebuild_enclosures(self._all_params)
        self._rebuild_param_rows(self._visible_params())
        self._rebuild_checkboxes(
            self.cond_holder, self.cond_vars, preview["conditions"])

        if preview["file_count"] == 0:
            self._set_status("No run CSVs found — check the folder structure.")
        else:
            self._set_status(f"Found {preview['file_count']} run files. Ready.")
        self._log("Scan complete.\n")

    def _rebuild_checkboxes(self, holder, var_store, names):
        for child in holder.winfo_children():
            child.destroy()
        var_store.clear()
        if not names:
            self._placeholder(holder, "none detected")
            return
        for i, name in enumerate(names):
            var = tk.BooleanVar(value=True)
            var_store[name] = var
            ttk.Checkbutton(holder, text=name, variable=var).grid(
                row=i // 3, column=i % 3, sticky="w", padx=(0, 12))

    def _rebuild_enclosures(self, params):
        """Build the EN# checkboxes from the parameter list. Toggling one
        re-filters which parameters are shown."""
        holder = self.enc_holder
        for child in holder.winfo_children():
            child.destroy()
        encs = sorted({e for e in (visualize.infer_enclosure(p) for p in params) if e})
        # Preserve prior selections where the enclosure still exists.
        prev = {k: v.get() for k, v in self.enc_vars.items()}
        self.enc_vars.clear()
        if not encs:
            self._placeholder(holder, "no EN# tags found")
            return
        for i, enc in enumerate(encs):
            var = tk.BooleanVar(value=prev.get(enc, True))
            var.trace_add("write", lambda *_: self._on_enclosure_change())
            self.enc_vars[enc] = var
            ttk.Checkbutton(holder, text=enc, variable=var).grid(
                row=0, column=i, sticky="w", padx=(0, 12))

    def _visible_params(self):
        """Parameters allowed by the current EN# selection (all if none set)."""
        selected = {e for e, v in self.enc_vars.items() if v.get()}
        if not selected:
            return list(self._all_params)
        out = []
        for p in self._all_params:
            enc = visualize.infer_enclosure(p)
            if not enc or enc in selected:
                out.append(p)
        return out

    def _on_enclosure_change(self):
        self._rebuild_param_rows(self._visible_params())

    def _rebuild_param_rows(self, names, per_row=3):
        """Parameter list wrapped row-major into ``per_row`` columns. Each cell
        holds the checkbox + Y-unit entry. The first ``per_row`` parameters
        fill row 1; the next group fills row 2; and so on. Check state and unit
        text are preserved across rebuilds via caches."""
        holder = self.param_holder
        # Cache current widget state so re-filtering doesn't lose user edits.
        for p, var in self.param_vars.items():
            self._param_check_cache[p] = bool(var.get())
        for p, var in self.param_unit_vars.items():
            self._param_unit_cache[p] = var.get()

        for child in holder.winfo_children():
            child.destroy()
        self.param_vars.clear()
        self.param_unit_vars.clear()
        if not names:
            self._placeholder(holder, "none detected")
            return

        # Header strip — one "Parameter / Y unit" pair above each grid column.
        for c in range(per_row):
            base = c * 2
            ttk.Label(holder, text="Parameter", foreground="#888888").grid(
                row=0, column=base, sticky="w", padx=(0, 4))
            ttk.Label(holder, text="Y unit", foreground="#888888").grid(
                row=0, column=base + 1, sticky="w", padx=(0, 16))

        for c in range(per_row):
            holder.columnconfigure(c * 2,     weight=1)
            holder.columnconfigure(c * 2 + 1, weight=0)

        for i, name in enumerate(names):
            r = (i // per_row) + 1           # data rows start at 1 (after header)
            c = i % per_row
            base = c * 2

            chk_var = tk.BooleanVar(value=self._param_check_cache.get(name, True))
            self.param_vars[name] = chk_var
            ttk.Checkbutton(holder, text=name, variable=chk_var).grid(
                row=r, column=base, sticky="w", padx=(0, 4), pady=1)

            unit_var = tk.StringVar(
                value=self._param_unit_cache.get(name, visualize.infer_unit(name)))
            self.param_unit_vars[name] = unit_var
            ttk.Entry(holder, textvariable=unit_var, width=5).grid(
                row=r, column=base + 1, sticky="w", padx=(0, 16), pady=1)

    # ── Run (background) ─────────────────────────────────────────────────────────

    def _start_run(self):
        if self._busy():
            return
        cfg = self._collect_and_validate()
        if cfg is None:
            return
        self.run_button.configure(state="disabled")
        self.progress_var.set(0)
        self._set_status("Working…")
        self.worker = threading.Thread(
            target=self._run_worker, args=(cfg,), daemon=True)
        self.worker.start()

    def _collect_and_validate(self):
        root_dir = self.input_dir_var.get().strip()
        out_dir  = self.output_dir_var.get().strip()
        name     = self.report_name_var.get().strip()

        if not root_dir or not os.path.isdir(root_dir):
            messagebox.showerror("Input folder", "Choose a valid input data folder.")
            return None
        if not out_dir:
            messagebox.showerror("Output folder", "Choose an output folder.")
            return None
        if not name:
            messagebox.showerror("Report name", "Enter a report name.")
            return None

        outputs = {
            "xlsx": self.out_xlsx_var.get(),
            "png":  self.out_png_var.get(),
            "csv":  self.out_csv_var.get(),
        }
        if not any(outputs.values()):
            messagebox.showwarning("No output", "Select at least one output format.")
            return None

        plots = [p for p, v in self.plot_vars.items() if v.get()]
        if outputs["png"] and not plots:
            messagebox.showwarning(
                "No plots", "Plot images selected but no plot types are ticked.")
            return None

        params = [p for p, v in self.param_vars.items() if v.get()]
        conds  = [c for c, v in self.cond_vars.items() if v.get()]
        if self.param_vars and not params:
            messagebox.showwarning("No parameters", "Select at least one parameter.")
            return None
        if self.cond_vars and not conds:
            messagebox.showwarning("No conditions", "Select at least one condition.")
            return None

        return {
            "root": root_dir,
            "out_dir": out_dir,
            "name": name,
            "outputs": outputs,
            "plots": plots,
            "params": set(params),
            "conds": set(conds),
            "units": {p: v.get().strip() for p, v in self.param_unit_vars.items()
                      if v.get().strip()},
            "img_format": self.img_format_var.get(),
            "dpi": int(self.dpi_var.get()),
            "theme": self.theme_var.get(),
            "temp_map": self._temp_map(),
            "boltzmann_opts": self._boltzmann_opts(),
        }

    def _temp_map(self):
        """Parse the COLD/ROOM/HOT temperature entries into {cond: °C}."""
        tmap = {}
        for cond, var in self.temp_vars.items():
            val = self._parse_float(var.get())
            if val is not None:
                tmap[cond] = val
        return tmap

    def _run_worker(self, cfg):
        try:
            self._do_run(cfg)
            self._post("done", None)
        except Exception:
            self._post("error", traceback.format_exc())

    def _do_run(self, cfg):
        log = lambda m: self._post("log", m)

        # Phase weighting for the overall progress bar.
        produce_plots = cfg["outputs"]["png"] and cfg["plots"]
        produce_xlsx  = cfg["outputs"]["xlsx"]
        produce_csv   = cfg["outputs"]["csv"]
        extract_span = 50.0
        rest = 100.0 - extract_span
        n_rest = sum([bool(produce_plots), bool(produce_xlsx), bool(produce_csv)]) or 1
        step = rest / n_rest
        base = [0.0]

        def phase_progress(span):
            start = base[0]
            def cb(done, total):
                frac = (done / total) if total else 1.0
                self._post("progress", start + span * frac)
            return cb

        # 1) Extract -------------------------------------------------------------
        log("=== Extracting ===")
        total_files = (self.last_preview or {}).get("file_count") or None
        if total_files is None:
            preview = extract.preview_scan(cfg["root"], log=lambda *_: None)
            total_files = preview["file_count"]
        records = extract.extract_records(
            cfg["root"], log=log,
            progress_callback=phase_progress(extract_span), total=total_files)
        base[0] = extract_span

        if not records:
            log("[ERROR] No run files found — nothing to do.")
            self._post("progress", 100)
            return

        raw_df = pd.DataFrame(records)
        df = visualize.enrich(raw_df)

        # Apply parameter / condition filters.
        if cfg["params"]:
            df = df[df["parameter"].isin(cfg["params"])]
        if cfg["conds"]:
            df = df[df["temp_cond"].isin(cfg["conds"])]
        log(f"Selected {len(df)} measurements after filtering.")
        if df.empty:
            log("[ERROR] No data left after filtering — adjust the filters.")
            self._post("progress", 100)
            return

        out_dir = Path(cfg["out_dir"])
        out_dir.mkdir(parents=True, exist_ok=True)
        name = cfg["name"]
        written = []

        # 2) CSV -----------------------------------------------------------------
        if produce_csv:
            log("\n=== Writing CSV ===")
            csv_path = out_dir / f"{name}.csv"
            df[BASE_COLS].to_csv(csv_path, index=False)
            log(f"  Saved → {csv_path}")
            written.append(str(csv_path))
            base[0] += step
            self._post("progress", base[0])

        # 3) Plots ---------------------------------------------------------------
        plot_dir = None
        if produce_plots:
            log("\n=== Generating plots ===")
            plot_dir = out_dir / f"{name}_plots"
            plot_dir.mkdir(parents=True, exist_ok=True)
            bopts = dict(cfg.get("boltzmann_opts") or {})
            # Merge per-parameter units (from the Parameters panel) so each
            # parameter's plot is labelled with its own unit.
            if cfg.get("units"):
                bopts["units"] = dict(cfg["units"])
            visualize.generate_plots(
                df, cfg["plots"], str(plot_dir),
                fmt=cfg["img_format"], dpi=cfg["dpi"],
                theme=cfg["theme"], log=log,
                progress_callback=phase_progress(step),
                boltzmann_opts=bopts, temp_map=cfg.get("temp_map"))
            written.append(str(plot_dir))
            base[0] += step
            self._post("progress", base[0])

        # 4) Excel ---------------------------------------------------------------
        if produce_xlsx:
            log("\n=== Building Excel workbook ===")
            xlsx_path = out_dir / f"{name}.xlsx"
            build.build_workbook(
                df, str(xlsx_path), log=log,
                progress_callback=phase_progress(step),
                temp_map=cfg.get("temp_map"))
            written.append(str(xlsx_path))
            base[0] += step
            self._post("progress", base[0])

        self._post("progress", 100)
        log("\n=== Done ===")
        for w in written:
            log(f"  {w}")

        # Open the interactive results window (one tab per generated plot).
        if produce_plots:
            self._post("results", {
                "df": df,
                "params": sorted(df["parameter"].unique()),
                "plots": cfg["plots"],
                "plot_dir": str(plot_dir),
                "fmt": cfg["img_format"],
                "dpi": cfg["dpi"],
                "theme": cfg["theme"],
                "temp_map": cfg.get("temp_map"),
                "units": dict(cfg.get("units") or {}),
                "boltzmann_opts": dict(cfg.get("boltzmann_opts") or {}),
            })

    # ── Queue plumbing ───────────────────────────────────────────────────────────

    def _post(self, kind, payload):
        """Called from worker threads only — never touches Tk directly."""
        self.msg_queue.put((kind, payload))

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    self.progress_var.set(max(0, min(100, payload)))
                elif kind == "status":
                    self._set_status(payload)
                elif kind == "preview":
                    self._apply_preview(payload)
                elif kind == "preview_image":
                    self._set_preview_image(*payload)
                elif kind == "preview_missing":
                    lbl = self._preview_labels.get(payload)
                    if lbl is not None:
                        lbl.configure(image="", text="(no preview)")
                elif kind == "results":
                    self._open_results_window(payload)
                elif kind == "done":
                    self._set_status("Done.")
                    self.run_button.configure(state="normal")
                elif kind == "error":
                    self._log("\n[ERROR]\n" + payload)
                    self._set_status("Error — see log.")
                    self.run_button.configure(state="normal")
                    messagebox.showerror(
                        "Error", "Something went wrong. See the log for details.")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    def _set_preview_image(self, key, path):
        """Load a freshly-rendered preview PNG, cache the source, and display
        it scaled to the current tab size."""
        lbl = self._preview_labels.get(key)
        if lbl is None:
            return
        try:
            src = Image.open(path).copy()
        except Exception:
            lbl.configure(image="", text="(preview unavailable)")
            return
        self._preview_src[key] = src
        self._preview_size[key] = None      # force a re-fit
        self._fit_preview(key)

    def _fit_preview(self, key):
        """Scale the cached source image to span the full preview width;
        height follows the source aspect ratio."""
        lbl = self._preview_labels.get(key)
        src = self._preview_src.get(key)
        if lbl is None or src is None or src.width <= 0:
            return
        w = max(lbl.winfo_width() - 8, 200)
        if self._preview_size.get(key) == w:
            return                          # already at this width — avoid churn
        self._preview_size[key] = w
        new_h = max(1, int(round(src.height * (w / src.width))))
        img = src.resize((w, new_h), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        self._preview_images[key] = photo   # keep ref
        lbl.configure(image=photo, text="")

    def _on_preview_configure(self, key, event):
        """Rescale the preview when its tab/label is resized."""
        if self._preview_src.get(key) is None:
            return
        last = self._preview_size.get(key)
        if last is not None and abs(event.width - 8 - last) < 4:
            return
        self._fit_preview(key)

    def _popout_preview(self, key):
        """Open a resizable window showing the full preview for this plot."""
        src = self._preview_src.get(key)
        if src is None:
            messagebox.showinfo("Preview", "No preview rendered yet.")
            return
        top = tk.Toplevel(self.root)
        top.title(f"{PLOT_LABELS.get(key, key)} — preview")
        top.geometry("960x600")
        lbl = ttk.Label(top, anchor="center", background="#ffffff")
        lbl.pack(fill="both", expand=True)
        state = {"photo": None, "size": None}

        def render(w, h):
            w, h = max(w - 12, 200), max(h - 12, 150)
            if state["size"] == (w, h):
                return
            state["size"] = (w, h)
            img = src.copy()
            img.thumbnail((w, h), Image.LANCZOS)
            state["photo"] = ImageTk.PhotoImage(img)
            lbl.configure(image=state["photo"])

        lbl.bind("<Configure>", lambda e: render(e.width, e.height))
        self._popouts.append(top)
        top.protocol("WM_DELETE_WINDOW",
                     lambda: (self._popouts.remove(top) if top in self._popouts
                              else None, top.destroy()))

    # ── Small helpers ────────────────────────────────────────────────────────────

    def _busy(self):
        return self.worker is not None and self.worker.is_alive()

    def _set_status(self, text):
        self.status_var.set(text)

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ── Results window ───────────────────────────────────────────────────────────

    def _open_results_window(self, ctx):
        try:
            ResultsWindow(self.root, ctx)
        except Exception:
            self._log("\n[ERROR opening results window]\n" + traceback.format_exc())


class ResultsWindow:
    """A separate window with one tab per generated plot image. Each tab shows
    the rendered plot plus the controls to customize it, and a Re-render button
    that overwrites the corresponding file in the report's plot folder."""

    def __init__(self, master, ctx):
        self.ctx = ctx
        self.df = ctx["df"]
        self.plot_dir = Path(ctx["plot_dir"])
        self.fmt = ctx["fmt"]
        self.dpi = ctx["dpi"]
        self.theme = ctx["theme"]
        self.temp_map = ctx.get("temp_map")
        self.units = ctx.get("units") or {}
        self.bopts = ctx.get("boltzmann_opts") or {}

        self.win = tk.Toplevel(master)
        self.win.title("Results — generated plots")
        self.win.geometry("1150x780")
        self.win.columnconfigure(0, weight=1)
        self.win.rowconfigure(1, weight=1)

        ttk.Label(self.win, padding=8, foreground="#555555",
                  text=("Each tab is one generated plot. Adjust its controls and "
                        "Re-render — changes overwrite the file in the report's "
                        "plot folder.")).grid(row=0, column=0, sticky="ew")

        self.nb = ttk.Notebook(self.win)
        self.nb.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

        # Per-tab state
        self._img_refs = {}     # tab_id -> PhotoImage
        self._src      = {}     # tab_id -> PIL.Image
        self._labels   = {}     # tab_id -> preview Label
        self._size     = {}     # tab_id -> last fitted width
        self._ctrls    = {}     # tab_id -> dict of control variables

        self._build_tabs()

    # ── Tab construction ─────────────────────────────────────────────────────────

    def _build_tabs(self):
        plots = self.ctx["plots"]
        if "boltzmann" in plots:
            for param in self.ctx["params"]:
                self._add_boltzmann_tab(param)
        labels = {"box": "Box-and-whisker", "trend": "Run trend",
                  "heatmap": "Heat map", "histogram": "Histogram"}
        for key in ("box", "trend", "heatmap", "histogram"):
            if key in plots:
                self._add_simple_tab(key, labels[key])

    def _preview_label(self, tab, tab_id):
        lbl = ttk.Label(tab, anchor="center", relief="sunken", background="#f0f0f0")
        lbl.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        tab.rowconfigure(1, weight=1)
        tab.columnconfigure(0, weight=1)
        lbl.bind("<Configure>", lambda e, tid=tab_id: self._fit(tid, e.width))
        self._labels[tab_id] = lbl
        return lbl

    def _add_boltzmann_tab(self, param):
        tab_id = f"boltzmann:{param}"
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text=param)
        tab.columnconfigure(0, weight=1)

        ctl = ttk.Frame(tab)
        ctl.grid(row=0, column=0, sticky="ew")

        title_var = tk.StringVar(value=param)
        unit_var  = tk.StringVar(value=self.units.get(param, visualize.infer_unit(param)))
        mode_var  = tk.StringVar(value=self.bopts.get("limit_mode", "sigma"))
        k_var     = tk.DoubleVar(value=self.bopts.get("sigma_k", 3.0))
        min_var   = tk.StringVar(value="")
        max_var   = tk.StringVar(value="")
        table_var = tk.BooleanVar(value=bool(self.bopts.get("show_table", True)))

        r1 = ttk.Frame(ctl); r1.grid(row=0, column=0, sticky="ew")
        ttk.Label(r1, text="Title:").grid(row=0, column=0, padx=(0, 4))
        ttk.Entry(r1, textvariable=title_var, width=22).grid(row=0, column=1, padx=(0, 12))
        ttk.Label(r1, text="Y unit:").grid(row=0, column=2, padx=(0, 4))
        ttk.Entry(r1, textvariable=unit_var, width=6).grid(row=0, column=3, padx=(0, 12))
        ttk.Checkbutton(r1, text="Value table", variable=table_var).grid(row=0, column=4)

        r2 = ttk.Frame(ctl); r2.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        ttk.Label(r2, text="Limits:").grid(row=0, column=0, padx=(0, 4))
        ttk.Combobox(r2, textvariable=mode_var, values=visualize.LIMIT_MODES,
                     width=8, state="readonly").grid(row=0, column=1, padx=(0, 12))
        ttk.Label(r2, text="k·σ:").grid(row=0, column=2, padx=(0, 4))
        ttk.Spinbox(r2, from_=0.5, to=10, increment=0.5, width=5,
                    textvariable=k_var).grid(row=0, column=3, padx=(0, 12))
        ttk.Label(r2, text="Min:").grid(row=0, column=4, padx=(0, 4))
        ttk.Entry(r2, textvariable=min_var, width=9).grid(row=0, column=5, padx=(0, 8))
        ttk.Label(r2, text="Max:").grid(row=0, column=6, padx=(0, 4))
        ttk.Entry(r2, textvariable=max_var, width=9).grid(row=0, column=7, padx=(0, 12))
        ttk.Button(r2, text="Re-render",
                   command=lambda p=param: self._render_boltzmann(p)).grid(row=0, column=8)

        self._ctrls[tab_id] = dict(title=title_var, unit=unit_var, mode=mode_var,
                                   k=k_var, min=min_var, max=max_var, table=table_var)
        self._preview_label(tab, tab_id)
        self._load_existing(tab_id, f"1_boltzmann_{param}")

    def _add_simple_tab(self, key, label):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text=label)
        tab.columnconfigure(0, weight=1)

        ctl = ttk.Frame(tab)
        ctl.grid(row=0, column=0, sticky="ew")
        title_var = tk.StringVar(value="")
        ttk.Label(ctl, text="Title:").grid(row=0, column=0, padx=(0, 4))
        ttk.Entry(ctl, textvariable=title_var, width=40).grid(row=0, column=1, padx=(0, 8))
        ttk.Label(ctl, text="(blank = default)", foreground="#888888").grid(row=0, column=2)
        ttk.Button(ctl, text="Re-render",
                   command=lambda k=key: self._render_simple(k)).grid(row=0, column=3, padx=(12, 0))

        self._ctrls[key] = dict(title=title_var)
        self._preview_label(tab, key)
        fname = {"box": "2_boxplot", "trend": "3_run_trend",
                 "heatmap": "4_heatmap", "histogram": "5_histogram"}[key]
        self._load_existing(key, fname)

    # ── Rendering ────────────────────────────────────────────────────────────────

    def _parse_float(self, s):
        try:
            return float(s)
        except (TypeError, ValueError):
            return None

    def _render_boltzmann(self, param):
        tab_id = f"boltzmann:{param}"
        c = self._ctrls[tab_id]
        lbl = self._labels[tab_id]
        lbl.configure(text="rendering…", image="")
        self.win.update_idletasks()
        mode = c["mode"].get()
        unit = c["unit"].get().strip()
        opts = dict(
            limit_mode=mode,
            sigma_k=self._parse_float(c["k"].get()) or 3.0,
            show_table=bool(c["table"].get()),
            show_title=True,
            title=c["title"].get().strip() or None,
            units={param: unit} if unit else None,
            temp_map=self.temp_map,
        )
        if mode == "fixed":
            opts["limits"] = {param: (self._parse_float(c["min"].get()),
                                      self._parse_float(c["max"].get()))}
        try:
            visualize.apply_base_style(self.theme)
            sub = self.df[self.df["parameter"] == param]
            visualize.plot_boltzmann(sub, str(self.plot_dir), self.fmt,
                                     dpi=self.dpi, theme=self.theme,
                                     log=lambda *_: None, **opts)
            self._load_existing(tab_id, f"1_boltzmann_{param}")
        except Exception:
            lbl.configure(text="render error:\n" + traceback.format_exc())

    def _render_simple(self, key):
        c = self._ctrls[key]
        lbl = self._labels[key]
        lbl.configure(text="rendering…", image="")
        self.win.update_idletasks()
        func = {"box": visualize.plot_boxplot, "trend": visualize.plot_run_trend,
                "heatmap": visualize.plot_heatmap,
                "histogram": visualize.plot_histograms}[key]
        fname = {"box": "2_boxplot", "trend": "3_run_trend",
                 "heatmap": "4_heatmap", "histogram": "5_histogram"}[key]
        title = c["title"].get().strip() or None
        try:
            visualize.apply_base_style(self.theme)
            func(self.df, str(self.plot_dir), self.fmt, dpi=self.dpi,
                 theme=self.theme, log=lambda *_: None, title=title)
            self._load_existing(key, fname)
        except Exception:
            lbl.configure(text="render error:\n" + traceback.format_exc())

    # ── Image display ────────────────────────────────────────────────────────────

    def _load_existing(self, tab_id, stem):
        path = self.plot_dir / f"{stem}.{self.fmt}"
        lbl = self._labels[tab_id]
        if not path.exists():
            lbl.configure(text="(not generated)", image="")
            self._src.pop(tab_id, None)
            return
        try:
            self._src[tab_id] = Image.open(path).copy()
        except Exception:
            lbl.configure(text="(image unavailable)", image="")
            return
        self._size[tab_id] = None
        self._fit(tab_id, lbl.winfo_width())

    def _fit(self, tab_id, width):
        src = self._src.get(tab_id)
        lbl = self._labels.get(tab_id)
        if src is None or lbl is None or src.width <= 0:
            return
        w = max(width - 8, 200)
        if self._size.get(tab_id) == w:
            return
        self._size[tab_id] = w
        new_h = max(1, int(round(src.height * (w / src.width))))
        img = src.resize((w, new_h), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        self._img_refs[tab_id] = photo
        lbl.configure(image=photo, text="")


def main():
    root = tk.Tk()
    ATEReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
