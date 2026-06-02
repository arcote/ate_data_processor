"""
visualize_ate_data.py
---------------------
Reads the output CSV from extract_ate_data.py and produces multiple
publication-quality plots for ATE test data review.

Usage:
    python visualize_ate_data.py --input results.csv
    python visualize_ate_data.py --input results.csv --output-dir ./plots --format pdf

Plots generated (all saved + shown interactively):
  1. Boltzmann scatter    — mirrors the reference format; one strip per temp condition
  2. Box-and-whisker      — distribution spread per parameter × temp condition
  3. Run trend            — value vs. run number per temp condition (drift/stability check)
  4. Heat map             — mean value grid: test iteration × temp condition
  5. Histogram overlay    — distribution shape per temp condition, all params together

Requirements:
    pip install pandas matplotlib seaborn scipy
"""

import argparse
import os
import sys
import re
import warnings

import matplotlib
# Default to the non-interactive Agg backend so this module is safe to import
# from a GUI/worker thread or a headless server. A standalone CLI run can opt
# back into an interactive backend via --show (handled in main()).
if not os.environ.get("MPLBACKEND"):
    matplotlib.use("Agg")

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

warnings.filterwarnings("ignore", category=FutureWarning)

# ── Temp-condition colour palette (matches reference plot aesthetic) ─────────
TEMP_PALETTE = {
    "COLD": "#3a7abf",   # medium blue  (matches reference left cluster)
    "ROOM": "#e040a0",   # vivid pink   (matches reference centre cluster)
    "HOT":  "#00a898",   # teal-green   (matches reference right cluster)
}
DEFAULT_COLOR = "#999999"

# ── Approximate temperature axis positions for Boltzmann plot ────────────────
TEMP_AXIS_MAP = {
    "COLD": -40,
    "ROOM":  25,
    "HOT":   85,
}

# Map column 4 dtype
VALUE_COL = "extracted_value"


# ── Themes ───────────────────────────────────────────────────────────────────
# A theme bundles every colour the plots use so the whole report can be
# switched between a light and a dark presentation. ``temp_palette`` is the
# per-condition series colour and falls back to ``default_color``.

THEMES = {
    "light": {
        "fig_face":     "#ffffff",
        "axes_face":    "#f7f8fc",
        "axes_edge":    "#cccccc",
        "text":         "#222222",
        "label":        "#222222",
        "title":        "#111111",
        "grid":         "#e0e0e0",
        "tick":         "#444444",
        "legend_face":  "#ffffff",
        "legend_edge":  "#cccccc",
        "median":       "#222222",
        "whisker":      "#8891aa",
        "heatmap_cmap": "YlOrRd",
        "heatmap_line": "#ffffff",
        "temp_palette": TEMP_PALETTE,
        "default_color": DEFAULT_COLOR,
    },
    "dark": {
        "fig_face":     "#0f1117",
        "axes_face":    "#1a1d28",
        "axes_edge":    "#3a3f4f",
        "text":         "#e6e6e6",
        "label":        "#dddddd",
        "title":        "#ffffff",
        "grid":         "#2a2e3a",
        "tick":         "#bbbbbb",
        "legend_face":  "#1a1d28",
        "legend_edge":  "#3a3f4f",
        "median":       "#ffffff",
        "whisker":      "#aab1c4",
        "heatmap_cmap": "magma",
        "heatmap_line": "#0f1117",
        "temp_palette": {"COLD": "#5fa8ff", "ROOM": "#ff6fc8", "HOT": "#34d4be"},
        "default_color": "#bbbbbb",
    },
}
DEFAULT_THEME = "light"


def get_theme(theme):
    """Resolve ``theme`` (a name or a custom dict) to a theme dict."""
    if isinstance(theme, dict):
        return theme
    if theme is None:
        return THEMES[DEFAULT_THEME]
    try:
        return THEMES[theme]
    except KeyError:
        raise ValueError(f"Unknown theme {theme!r}. "
                         f"Choices: {sorted(THEMES)}")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def detect_temp_condition(temp_folder: str) -> str:
    """Return 'COLD', 'ROOM', or 'HOT' from a folder name like 'Idss_EN5_COLD'."""
    upper = temp_folder.upper()
    for key in ("COLD", "HOT", "ROOM"):
        if key in upper:
            return key
    return "ROOM"


def detect_parameter(temp_folder: str) -> str:
    """
    Extract parameter name from folder like 'Idss_EN5_COLD' → 'Idss_EN5'.
    Strips trailing _COLD / _HOT / _ROOM suffix.
    """
    return re.sub(r"_(COLD|HOT|ROOM)$", "", temp_folder, flags=re.IGNORECASE)


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add the derived columns the plots rely on (temp_cond, parameter, temp_x,
    color) and coerce numerics. Idempotent — safe to call on a frame that has
    already been enriched. Operates on a copy.
    """
    required = {"test_folder", "temp_folder", "run_number", VALUE_COL}
    missing  = required - set(df.columns)
    if missing:
        raise ValueError(f"Input is missing columns: {missing}\n"
                         f"       Expected columns: {required}")

    df = df.copy()
    df[VALUE_COL]     = pd.to_numeric(df[VALUE_COL], errors="coerce")
    df["run_number"]  = pd.to_numeric(df["run_number"], errors="coerce")
    df["temp_cond"]   = df["temp_folder"].apply(detect_temp_condition)
    df["parameter"]   = df["temp_folder"].apply(detect_parameter)
    df["temp_x"]      = df["temp_cond"].map(TEMP_AXIS_MAP)
    df["color"]       = df["temp_cond"].map(TEMP_PALETTE).fillna(DEFAULT_COLOR)
    df.dropna(subset=[VALUE_COL], inplace=True)

    return df


def load_and_enrich(csv_path: str) -> pd.DataFrame:
    """Read the extraction CSV and enrich it. Exits the process on bad input."""
    try:
        df = pd.read_csv(csv_path)
        return enrich(df)
    except ValueError as exc:
        sys.exit(f"[ERROR] {exc}")


def save_fig(fig, out_dir: str, name: str, fmt: str, dpi: int = 150,
             facecolor: str = "#ffffff", log=print):
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{name}.{fmt}")
    fig.savefig(path, dpi=dpi, bbox_inches="tight", facecolor=facecolor)
    log(f"  Saved → {path}")
    return path


def apply_base_style(theme=None):
    """Update matplotlib rcParams from a theme dict (or a theme name)."""
    t = get_theme(theme)
    plt.rcParams.update({
        "figure.facecolor":  t["fig_face"],
        "axes.facecolor":    t["axes_face"],
        "axes.edgecolor":    t["axes_edge"],
        "axes.labelcolor":   t["label"],
        "axes.titlecolor":   t["title"],
        "axes.grid":         True,
        "grid.color":        t["grid"],
        "grid.linewidth":    0.7,
        "xtick.color":       t["tick"],
        "ytick.color":       t["tick"],
        "text.color":        t["text"],
        "legend.facecolor":  t["legend_face"],
        "legend.edgecolor":  t["legend_edge"],
        "legend.labelcolor": t["text"],
        "font.family":       "monospace",
        "font.size":         9,
    })


# ─────────────────────────────────────────────────────────────────────────────
# Plot 1 — Boltzmann Scatter (reference format)
# ─────────────────────────────────────────────────────────────────────────────

def plot_boltzmann(df: pd.DataFrame, out_dir: str, fmt: str, dpi: int = 150,
                   theme=None, log=print):
    """One scatter strip per temp condition, x-axis is temperature (°C)."""
    t = get_theme(theme)
    palette = t["temp_palette"]
    params = df["parameter"].unique()

    for param in params:
        sub = df[df["parameter"] == param].copy()

        fig, ax = plt.subplots(figsize=(10, 4.5))
        fig.patch.set_facecolor(t["fig_face"])

        jitter = 1.2   # x-axis spread for visual separation of overlapping dots

        for cond, grp in sub.groupby("temp_cond"):
            x_base = TEMP_AXIS_MAP.get(cond, 25)
            # Light horizontal jitter so overlapping points separate
            x_jitter = x_base + (grp["run_number"].values - grp["run_number"].mean()) / max(grp["run_number"].std() or 1, 1) * jitter

            ax.scatter(
                x_jitter,
                grp[VALUE_COL],
                c=palette.get(cond, t["default_color"]),
                alpha=0.75,
                s=28,
                zorder=3,
                label=cond,
            )

        ax.set_xlabel("Temperature (°C)", fontsize=10)
        ax.set_ylabel(f"{param}\n(extracted value)", fontsize=10)
        ax.set_title(f"Boltzmann Scatter — {param}", fontsize=12, pad=10)

        # X-axis: show only the three temp positions
        used_temps = sorted(sub["temp_cond"].unique(), key=lambda t: TEMP_AXIS_MAP.get(t, 0))
        ax.set_xticks([TEMP_AXIS_MAP[t] for t in used_temps])
        ax.set_xticklabels([f"{TEMP_AXIS_MAP[t]} °C\n({t})" for t in used_temps])
        ax.set_xlim(min(TEMP_AXIS_MAP[t] for t in used_temps) - 15,
                    max(TEMP_AXIS_MAP[t] for t in used_temps) + 15)

        handles, labels = ax.get_legend_handles_labels()
        by_label = dict(zip(labels, handles))
        ax.legend(by_label.values(), by_label.keys(), loc="upper left")

        fig.tight_layout()
        save_fig(fig, out_dir, f"1_boltzmann_{param}", fmt, dpi=dpi,
                 facecolor=t["fig_face"], log=log)
        plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# Plot 2 — Box-and-Whisker
# ─────────────────────────────────────────────────────────────────────────────

def plot_boxplot(df: pd.DataFrame, out_dir: str, fmt: str, dpi: int = 150,
                 theme=None, log=print):
    """Distribution spread per parameter × temperature condition."""
    t = get_theme(theme)
    palette = t["temp_palette"]
    params = df["parameter"].unique()
    n = len(params)
    ncols = min(3, n)
    nrows = (n + ncols - 1) // ncols

    fig, axes = plt.subplots(nrows, ncols, figsize=(5 * ncols, 4.5 * nrows), squeeze=False)
    fig.patch.set_facecolor(t["fig_face"])
    fig.suptitle("Box-and-Whisker: Value Distribution per Condition", fontsize=13, y=1.01)

    for idx, param in enumerate(params):
        ax = axes[idx // ncols][idx % ncols]
        sub = df[df["parameter"] == param]

        order  = [c for c in ("COLD", "ROOM", "HOT") if c in sub["temp_cond"].unique()]
        colors = [palette.get(c, t["default_color"]) for c in order]

        bp = ax.boxplot(
            [sub[sub["temp_cond"] == c][VALUE_COL].values for c in order],
            patch_artist=True,
            notch=False,
            widths=0.45,
            medianprops=dict(color=t["median"], linewidth=1.8),
            whiskerprops=dict(color=t["whisker"]),
            capprops=dict(color=t["whisker"]),
            flierprops=dict(marker="o", markersize=3, alpha=0.5),
        )
        for patch, color in zip(bp["boxes"], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.7)
        for flier, color in zip(bp["fliers"], colors):
            flier.set_markerfacecolor(color)

        ax.set_xticklabels(order)
        ax.set_title(param, fontsize=9)
        ax.set_ylabel("Value", fontsize=8)
        ax.set_xlabel("Temp Condition", fontsize=8)

    # Hide unused axes
    for idx in range(n, nrows * ncols):
        axes[idx // ncols][idx % ncols].set_visible(False)

    fig.tight_layout()
    save_fig(fig, out_dir, "2_boxplot", fmt, dpi=dpi,
             facecolor=t["fig_face"], log=log)
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# Plot 3 — Run Trend (stability / drift check)
# ─────────────────────────────────────────────────────────────────────────────

def plot_run_trend(df: pd.DataFrame, out_dir: str, fmt: str, dpi: int = 150,
                   theme=None, log=print):
    """Value vs. run number — reveals drift, outliers, repeatability issues."""
    t = get_theme(theme)
    palette = t["temp_palette"]
    params = df["parameter"].unique()
    n = len(params)
    ncols = min(2, n)
    nrows = (n + ncols - 1) // ncols

    fig, axes = plt.subplots(nrows, ncols, figsize=(7 * ncols, 4 * nrows), squeeze=False)
    fig.patch.set_facecolor(t["fig_face"])
    fig.suptitle("Run Trend: Value vs. Run Number (Drift / Stability)", fontsize=13, y=1.01)

    for idx, param in enumerate(params):
        ax = axes[idx // ncols][idx % ncols]
        sub = df[df["parameter"] == param].sort_values("run_number")

        for cond, grp in sub.groupby("temp_cond"):
            grp_s = grp.sort_values("run_number")
            color = palette.get(cond, t["default_color"])
            ax.plot(
                grp_s["run_number"], grp_s[VALUE_COL],
                marker="o", ms=5, lw=1.4,
                color=color, label=cond, alpha=0.85
            )
            # Rolling mean overlay
            if len(grp_s) >= 4:
                rolling = grp_s[VALUE_COL].rolling(3, center=True).mean()
                ax.plot(grp_s["run_number"], rolling,
                        lw=2.2, ls="--", color=color, alpha=0.45)

        ax.set_title(param, fontsize=9)
        ax.set_xlabel("Run #", fontsize=8)
        ax.set_ylabel("Value", fontsize=8)
        ax.legend(fontsize=7, loc="upper right")
        ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    for idx in range(n, nrows * ncols):
        axes[idx // ncols][idx % ncols].set_visible(False)

    fig.tight_layout()
    save_fig(fig, out_dir, "3_run_trend", fmt, dpi=dpi,
             facecolor=t["fig_face"], log=log)
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# Plot 4 — Heat Map (mean value grid)
# ─────────────────────────────────────────────────────────────────────────────

def plot_heatmap(df: pd.DataFrame, out_dir: str, fmt: str, dpi: int = 150,
                 theme=None, log=print):
    """
    Mean extracted value as a colour grid: rows = parameter×condition,
    columns = test iteration (3rd/4th/5th test).
    Good for spotting lot-to-lot or iteration-to-iteration shifts.
    """
    t = get_theme(theme)
    df = df.copy()
    df["param_cond"] = df["parameter"] + " | " + df["temp_cond"]
    pivot = df.pivot_table(
        index="param_cond",
        columns="test_folder",
        values=VALUE_COL,
        aggfunc="mean"
    )

    if pivot.empty:
        log("  [SKIP] Heatmap: not enough variation across test folders.")
        return

    fig, ax = plt.subplots(figsize=(max(6, len(pivot.columns) * 2.2), max(4, len(pivot) * 0.55 + 1.5)))
    fig.patch.set_facecolor(t["fig_face"])

    sns.heatmap(
        pivot,
        ax=ax,
        cmap=t["heatmap_cmap"],
        annot=True,
        fmt=".3g",
        linewidths=0.4,
        linecolor=t["heatmap_line"],
        cbar_kws={"shrink": 0.7, "label": "Mean Value"},
    )

    ax.set_title("Heat Map: Mean Value — Parameter × Condition vs. Test Iteration", fontsize=11, pad=12)
    ax.set_xlabel("Test Iteration", fontsize=9)
    ax.set_ylabel("Parameter | Condition", fontsize=9)
    ax.tick_params(axis="x", rotation=30)
    ax.tick_params(axis="y", rotation=0)

    fig.tight_layout()
    save_fig(fig, out_dir, "4_heatmap", fmt, dpi=dpi,
             facecolor=t["fig_face"], log=log)
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# Plot 5 — Histogram Overlay
# ─────────────────────────────────────────────────────────────────────────────

def plot_histograms(df: pd.DataFrame, out_dir: str, fmt: str, dpi: int = 150,
                    theme=None, log=print):
    """Overlaid histograms per temp condition — shows distribution shape."""
    t = get_theme(theme)
    palette = t["temp_palette"]
    params = df["parameter"].unique()
    n = len(params)
    ncols = min(2, n)
    nrows = (n + ncols - 1) // ncols

    fig, axes = plt.subplots(nrows, ncols, figsize=(6.5 * ncols, 4 * nrows), squeeze=False)
    fig.patch.set_facecolor(t["fig_face"])
    fig.suptitle("Histogram Overlay: Value Distribution per Temp Condition", fontsize=13, y=1.01)

    for idx, param in enumerate(params):
        ax = axes[idx // ncols][idx % ncols]
        sub = df[df["parameter"] == param]
        bins = min(20, max(5, len(sub) // 4))

        for cond in ("COLD", "ROOM", "HOT"):
            grp = sub[sub["temp_cond"] == cond][VALUE_COL].dropna()
            if grp.empty:
                continue
            color = palette.get(cond, t["default_color"])
            ax.hist(
                grp, bins=bins,
                color=color, alpha=0.55,
                edgecolor=color, linewidth=0.5,
                label=f"{cond} (n={len(grp)})"
            )
            # Mean line
            ax.axvline(grp.mean(), color=color, lw=1.8, ls="--", alpha=0.9)

        ax.set_title(param, fontsize=9)
        ax.set_xlabel("Value", fontsize=8)
        ax.set_ylabel("Count", fontsize=8)
        ax.legend(fontsize=7)

    for idx in range(n, nrows * ncols):
        axes[idx // ncols][idx % ncols].set_visible(False)

    fig.tight_layout()
    save_fig(fig, out_dir, "5_histogram", fmt, dpi=dpi,
             facecolor=t["fig_face"], log=log)
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# Dispatcher
# ─────────────────────────────────────────────────────────────────────────────

ALL_PLOTS = ["boltzmann", "box", "trend", "heatmap", "histogram"]

_PLOT_FUNCS = {
    "boltzmann": plot_boltzmann,
    "box":       plot_boxplot,
    "trend":     plot_run_trend,
    "heatmap":   plot_heatmap,
    "histogram": plot_histograms,
}


def generate_plots(df, plots, out_dir, fmt="png", dpi=150, theme=None,
                   log=print, progress_callback=None):
    """
    Generate the requested plots from an enriched DataFrame.

    plots: iterable of names from ALL_PLOTS.
    theme: a name in THEMES (e.g. "light", "dark") or a custom theme dict.
    Applies the base style, then runs each selected plot. progress_callback(
    done, total) fires after each plot type. Returns the list of plots run.
    """
    t = get_theme(theme)
    apply_base_style(t)
    selected = [p for p in ALL_PLOTS if p in set(plots)]
    total = len(selected)
    for i, name in enumerate(selected, 1):
        log(f"[{i}/{total}] {name}...")
        _PLOT_FUNCS[name](df, out_dir, fmt, dpi=dpi, theme=t, log=log)
        if progress_callback is not None:
            progress_callback(i, total)
    return selected


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Visualize ATE extracted data from extract_ate_data.py output."
    )
    parser.add_argument(
        "--input", default="results.csv",
        help="Path to the extraction output CSV (default: results.csv)"
    )
    parser.add_argument(
        "--output-dir", default="plots",
        help="Directory to save plot files (default: ./plots)"
    )
    parser.add_argument(
        "--format", default="png", choices=["png", "pdf", "svg"],
        help="Output file format (default: png)"
    )
    parser.add_argument(
        "--dpi", type=int, default=150,
        help="Output resolution for raster formats (default: 150)"
    )
    parser.add_argument(
        "--theme", default=DEFAULT_THEME, choices=sorted(THEMES),
        help=f"Color theme for the plots (default: {DEFAULT_THEME})"
    )
    parser.add_argument(
        "--plots", default="all",
        help=(
            "Comma-separated list of plots to generate: "
            "boltzmann,box,trend,heatmap,histogram  (default: all)"
        )
    )
    args = parser.parse_args()

    print(f"Loading: {args.input}")
    df = load_and_enrich(args.input)
    print(f"  {len(df)} rows | {df['parameter'].nunique()} parameters | "
          f"{df['temp_cond'].nunique()} temp conditions | "
          f"{df['test_folder'].nunique()} test iterations\n")

    requested = (
        ALL_PLOTS
        if args.plots.lower() == "all"
        else [p.strip().lower() for p in args.plots.split(",")]
    )

    generate_plots(df, requested, args.output_dir, args.format,
                   dpi=args.dpi, theme=args.theme)

    print("\nAll done.")


if __name__ == "__main__":
    main()
