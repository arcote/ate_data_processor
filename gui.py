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
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

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
IMAGE_FORMATS = ["png", "pdf", "svg"]
BASE_COLS = ["test_folder", "temp_folder", "run_number", "extracted_value"]


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

        self.param_vars = {}   # parameter name -> BooleanVar
        self.cond_vars  = {}   # condition name -> BooleanVar
        self.plot_vars  = {p: tk.BooleanVar(value=True) for p in PLOT_LABELS}

        self.out_xlsx_var = tk.BooleanVar(value=True)
        self.out_png_var  = tk.BooleanVar(value=True)
        self.out_csv_var  = tk.BooleanVar(value=True)

        self.img_format_var = tk.StringVar(value="png")
        self.dpi_var        = tk.IntVar(value=150)

        self._build_ui()
        self.root.after(100, self._drain_queue)

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        outer = ttk.Frame(root, padding=10)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        # The log row should absorb extra vertical space.
        outer.rowconfigure(6, weight=1)

        self._build_input_row(outer, row=0)
        self._build_preview(outer, row=1)
        self._build_filters(outer, row=2)
        self._build_plots(outer, row=3)
        self._build_output(outer, row=4)
        self._build_run_row(outer, row=5)
        self._build_log(outer, row=6)

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

    def _build_filters(self, parent, row):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        pframe = ttk.LabelFrame(frame, text="Parameters", padding=8)
        pframe.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        cframe = ttk.LabelFrame(frame, text="Conditions", padding=8)
        cframe.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        self.param_holder = ttk.Frame(pframe)
        self.param_holder.grid(row=0, column=0, sticky="w")
        self.cond_holder = ttk.Frame(cframe)
        self.cond_holder.grid(row=0, column=0, sticky="w")

        self._placeholder(self.param_holder, "scan a folder to list parameters")
        self._placeholder(self.cond_holder, "scan a folder to list conditions")

    def _build_plots(self, parent, row):
        frame = ttk.LabelFrame(parent, text="Plots", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))

        toggles = ttk.Frame(frame)
        toggles.grid(row=0, column=0, sticky="w")
        for i, (key, label) in enumerate(PLOT_LABELS.items()):
            ttk.Checkbutton(toggles, text=label, variable=self.plot_vars[key]).grid(
                row=0, column=i, sticky="w", padx=(0, 12))

        sub = ttk.Frame(frame)
        sub.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(sub, text="Image format:").grid(row=0, column=0, padx=(0, 4))
        ttk.Combobox(sub, textvariable=self.img_format_var, values=IMAGE_FORMATS,
                     width=6, state="readonly").grid(row=0, column=1, padx=(0, 16))
        ttk.Label(sub, text="DPI:").grid(row=0, column=2, padx=(0, 4))
        ttk.Spinbox(sub, from_=72, to=600, increment=1, width=6,
                    textvariable=self.dpi_var).grid(row=0, column=3)

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

        self._rebuild_checkboxes(
            self.param_holder, self.param_vars, preview["parameters"])
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
            "img_format": self.img_format_var.get(),
            "dpi": int(self.dpi_var.get()),
        }

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
        if produce_plots:
            log("\n=== Generating plots ===")
            plot_dir = out_dir / f"{name}_plots"
            plot_dir.mkdir(parents=True, exist_ok=True)
            visualize.generate_plots(
                df, cfg["plots"], str(plot_dir),
                fmt=cfg["img_format"], dpi=cfg["dpi"], log=log,
                progress_callback=phase_progress(step))
            written.append(str(plot_dir))
            base[0] += step
            self._post("progress", base[0])

        # 4) Excel ---------------------------------------------------------------
        if produce_xlsx:
            log("\n=== Building Excel workbook ===")
            xlsx_path = out_dir / f"{name}.xlsx"
            build.build_workbook(
                df, str(xlsx_path), log=log,
                progress_callback=phase_progress(step))
            written.append(str(xlsx_path))
            base[0] += step
            self._post("progress", base[0])

        self._post("progress", 100)
        log("\n=== Done ===")
        for w in written:
            log(f"  {w}")

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


def main():
    root = tk.Tk()
    ATEReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
