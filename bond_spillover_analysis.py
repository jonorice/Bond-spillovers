"""
Euro area bond yield spillovers: Python recreation of R charts

This script recreates Charts 01-04 from the R version using only the yields data.
Charts 05-13 require external spillover metrics not available in this repo.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# -----------------------------
# Paths and settings
# -----------------------------
base_dir = Path("/home/user/Bond-spillovers")
out_dir = base_dir / "python_charts"
data_dir = out_dir / "chart_data"
out_dir.mkdir(exist_ok=True)
data_dir.mkdir(exist_ok=True)

in_yields_file = base_dir / "nav_test_clean.csv"

# Groups
outside_names = ["US", "JP", "UK"]
ea_names = ["DE", "FR", "IT", "ES"]

# Chart date windows
recent_start = pd.Timestamp("2014-01-01")
long_start = pd.Timestamp("1995-01-01")

# Smoothing windows (trading days)
k_yields_long = 63
k_yields_short = 21

# -----------------------------
# Prescribed colour scheme (fixed per country throughout)
# -----------------------------
country_cols = {
    "US": "#003299",
    "JP": "#FFB400",
    "UK": "#FF4B00",
    "DE": "#65B800",
    "FR": "#00B1EA",
    "IT": "#007816",
    "ES": "#8139C6"
}

# -----------------------------
# Events: numbered markers + keyed labels
# -----------------------------
events = pd.DataFrame({
    "id": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
    "date": pd.to_datetime([
        "2012-07-26",  # Draghi
        "2015-01-22",  # ECB QE announced
        "2016-06-23",  # Brexit referendum
        "2016-09-21",  # BoJ introduces YCC
        "2018-05-29",  # Italy political crisis
        "2019-09-12",  # ECB restarts QE
        "2020-03-18",  # ECB PEPP
        "2022-02-24",  # Ukraine invasion
        "2022-03-16",  # Fed hiking cycle
        "2022-07-21",  # ECB lift-off + TPI
        "2022-09-23",  # UK mini-budget
        "2022-12-20",  # BoJ widens YCC band
        "2024-03-19",  # BoJ ends NIRP/YCC
        "2024-06-06"   # ECB begins cutting
    ]),
    "label": [
        "ECB Draghi 'whatever it takes'",
        "ECB announces QE (PSPP)",
        "UK Brexit referendum",
        "BoJ introduces YCC",
        "Italy political crisis, BTP sell-off",
        "ECB restarts QE and tiering",
        "ECB launches PEPP",
        "Russia invades Ukraine",
        "Fed hiking cycle begins",
        "ECB lift-off and TPI",
        "UK mini-budget, gilt crisis",
        "BoJ widens YCC band",
        "BoJ ends NIRP and YCC",
        "ECB begins cutting cycle"
    ]
})

events.to_csv(data_dir / "events_key.csv", index=False)

# -----------------------------
# Helper functions
# -----------------------------
def rollmean_partial(x, k):
    """Centered rolling mean with partial windows at edges"""
    return x.rolling(window=k, center=True, min_periods=1).mean()

def calc_ylim(data_series, lower0=False, pad=0.22):
    """Calculate y-axis limits with padding"""
    all_vals = pd.concat([s for s in data_series if s is not None])
    all_vals = all_vals[np.isfinite(all_vals)]
    if len(all_vals) == 0:
        return (0, 1)
    r_min, r_max = all_vals.min(), all_vals.max()
    if r_min == r_max:
        r_min, r_max = r_min - 1, r_max + 1
    p = pad * (r_max - r_min)
    lo = 0 if lower0 else r_min - p
    hi = r_max + p
    return (lo, hi)

def calc_ylim_sym(data_series, pad=0.25):
    """Calculate symmetric y-axis limits"""
    all_vals = pd.concat([s for s in data_series if s is not None])
    all_vals = all_vals[np.isfinite(all_vals)]
    if len(all_vals) == 0:
        return (-1, 1)
    m = np.abs(all_vals).max()
    if not np.isfinite(m) or m == 0:
        m = 1
    m = m * (1 + pad)
    return (-m, m)

def filter_events(date_min, date_max, include_ids=None):
    """Filter events within date range"""
    mask = (events["date"] >= date_min) & (events["date"] <= date_max)
    ev = events[mask].copy()
    if include_ids is not None:
        ev = ev[ev["id"].isin(include_ids)]
    return ev

def add_events(ax, ev, ylim):
    """Add event markers to plot"""
    if len(ev) == 0:
        return
    for _, row in ev.iterrows():
        ax.axvline(row["date"], color="#B3B3B3", linestyle=":", linewidth=1, alpha=0.7)

    y_top = ylim[1] - 0.05 * (ylim[1] - ylim[0])
    stagger = [0, -0.03, -0.06] * (len(ev) // 3 + 1)

    for i, (_, row) in enumerate(ev.iterrows()):
        y_pos = y_top + stagger[i] * (ylim[1] - ylim[0])
        ax.text(row["date"], y_pos, str(row["id"]), fontsize=8, color="#666666",
                ha="center", va="bottom")

def event_key_text(ev):
    """Generate event key text for footer"""
    if len(ev) == 0:
        return "Events: none shown."
    parts = [f"{r['id']}={r['label']}" for _, r in ev.iterrows()]
    return "Events: " + "; ".join(parts) + "."

def wrap_text(text, width=100):
    """Simple text wrapping"""
    import textwrap
    return "\n".join(textwrap.wrap(text, width=width))

def zscore(x):
    """Standardize a series"""
    return (x - x.mean()) / x.std()

def plot_with_footer(out_png, main_plot_fn, footer_lines, figsize=(12, 7.5)):
    """Create plot with footer panel"""
    fig = plt.figure(figsize=figsize)

    # Main plot area
    ax_main = fig.add_axes([0.08, 0.22, 0.88, 0.70])
    main_plot_fn(ax_main)

    # Footer area
    ax_footer = fig.add_axes([0.08, 0.02, 0.88, 0.16])
    ax_footer.axis("off")
    footer_text = "\n".join(footer_lines)
    ax_footer.text(0, 1, footer_text, transform=ax_footer.transAxes,
                   fontsize=9, verticalalignment="top", family="monospace",
                   wrap=True)

    plt.savefig(out_png, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"Saved: {out_png}")

# -----------------------------
# Load yields
# -----------------------------
print("Loading yields data...")
y_raw = pd.read_csv(in_yields_file)

# Parse dates
y_raw["Date"] = pd.to_datetime(y_raw["Date"])
y_raw = y_raw.sort_values("Date").drop_duplicates(subset=["Date"], keep="last")

# Create yields dataframe
yields = y_raw[["Date"] + outside_names + ea_names].copy()
yields = yields.set_index("Date")

# Convert to numeric
for col in yields.columns:
    yields[col] = pd.to_numeric(yields[col], errors="coerce")

print(f"Loaded {len(yields)} observations from {yields.index.min()} to {yields.index.max()}")

# Derived spreads (basis points) and dispersion
yields["IT_DE_spread_bp"] = 100 * (yields["IT"] - yields["DE"])
yields["ES_DE_spread_bp"] = 100 * (yields["ES"] - yields["DE"])
yields["FR_DE_spread_bp"] = 100 * (yields["FR"] - yields["DE"])
yields["EA_dispersion_sd"] = yields[ea_names].std(axis=1)

# Smoothed yields
for nm in outside_names + ea_names:
    yields[f"{nm}_sm63"] = rollmean_partial(yields[nm], k_yields_long)
    yields[f"{nm}_sm21"] = rollmean_partial(yields[nm], k_yields_short)

yields["IT_DE_spread_bp_sm63"] = rollmean_partial(yields["IT_DE_spread_bp"], k_yields_long)
yields["ES_DE_spread_bp_sm63"] = rollmean_partial(yields["ES_DE_spread_bp"], k_yields_long)
yields["FR_DE_spread_bp_sm63"] = rollmean_partial(yields["FR_DE_spread_bp"], k_yields_long)
yields["IT_DE_spread_bp_sm21"] = rollmean_partial(yields["IT_DE_spread_bp"], k_yields_short)
yields["EA_dispersion_sd_sm21"] = rollmean_partial(yields["EA_dispersion_sd"], k_yields_short)

# Reset index for plotting
yields = yields.reset_index()

# ============================================================
# CHART 01: Long run yields (US, UK, JP, DE, IT)
# ============================================================
print("\nGenerating Chart 01: Long run yields...")
chart01 = yields[yields["Date"] >= long_start][["Date", "US_sm63", "UK_sm63", "JP_sm63", "DE_sm63", "IT_sm63"]].copy()
chart01.to_csv(data_dir / "chart01_longrun_yields_sm63.csv", index=False)

def plot_chart01(ax):
    cols_to_plot = ["US_sm63", "UK_sm63", "JP_sm63", "DE_sm63", "IT_sm63"]
    labels = ["US", "UK", "JP", "DE", "IT"]

    y_lim = calc_ylim([chart01[c] for c in cols_to_plot], pad=0.22)

    for col, label in zip(cols_to_plot, labels):
        country = label
        ax.plot(chart01["Date"], chart01[col], linewidth=2.2,
                color=country_cols[country], label=label)

    ax.set_ylim(y_lim)
    ax.set_xlabel("Date")
    ax.set_ylabel("10 year yield (percent)")
    ax.set_title("Long run 10 year yields (smoothed)", fontsize=12, fontweight="bold")
    ax.grid(True, color="#E6E6E6", linestyle="-", linewidth=0.5)
    ax.legend(loc="upper right", frameon=False)

    ev = filter_events(chart01["Date"].min(), chart01["Date"].max(),
                       include_ids=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
    add_events(ax, ev, y_lim)

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_major_locator(mdates.YearLocator(5))

ev01 = filter_events(chart01["Date"].min(), chart01["Date"].max(),
                     include_ids=[2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
plot_with_footer(
    out_dir / "chart01_longrun_yields_sm63.png",
    plot_chart01,
    [
        "What it shows: long-run rate regimes across major sovereign markets.",
        f"Method: daily 10 year yields smoothed using a {k_yields_long}-day centred moving average.",
        event_key_text(ev01)
    ]
)

# ============================================================
# CHART 02: Long run EA spreads to Germany (basis points)
# ============================================================
print("Generating Chart 02: EA spreads to Germany...")
chart02 = yields[yields["Date"] >= long_start][["Date", "IT_DE_spread_bp_sm63", "ES_DE_spread_bp_sm63", "FR_DE_spread_bp_sm63"]].copy()
chart02.to_csv(data_dir / "chart02_EA_spreads_to_DE_bp_sm63.csv", index=False)

def plot_chart02(ax):
    y_lim = calc_ylim([chart02["IT_DE_spread_bp_sm63"], chart02["ES_DE_spread_bp_sm63"],
                       chart02["FR_DE_spread_bp_sm63"]], lower0=True, pad=0.22)

    ax.plot(chart02["Date"], chart02["IT_DE_spread_bp_sm63"], linewidth=2.2,
            color=country_cols["IT"], label="IT minus DE")
    ax.plot(chart02["Date"], chart02["ES_DE_spread_bp_sm63"], linewidth=2.2,
            color=country_cols["ES"], label="ES minus DE")
    ax.plot(chart02["Date"], chart02["FR_DE_spread_bp_sm63"], linewidth=2.2,
            color=country_cols["FR"], label="FR minus DE")

    ax.set_ylim(y_lim)
    ax.set_xlabel("Date")
    ax.set_ylabel("Spread to Germany (basis points)")
    ax.set_title("Euro area spreads to Germany (smoothed)", fontsize=12, fontweight="bold")
    ax.grid(True, color="#E6E6E6", linestyle="-", linewidth=0.5)
    ax.legend(loc="upper right", frameon=False)

    ev = filter_events(chart02["Date"].min(), chart02["Date"].max(),
                       include_ids=[2, 5, 6, 7, 8, 10, 11, 14])
    add_events(ax, ev, y_lim)

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_major_locator(mdates.YearLocator(5))

ev02 = filter_events(chart02["Date"].min(), chart02["Date"].max(),
                     include_ids=[2, 5, 6, 7, 8, 10, 11, 14])
plot_with_footer(
    out_dir / "chart02_EA_spreads_to_DE_bp_sm63.png",
    plot_chart02,
    [
        "What it shows: EA fragmentation episodes through periphery and France spreads to Germany.",
        f"Method: spreads computed from daily 10 year yields (bp) and smoothed with a {k_yields_long}-day centred moving average.",
        event_key_text(ev02)
    ]
)

# ============================================================
# CHART 03: Recent yields: DE and IT versus US, UK, JP
# ============================================================
print("Generating Chart 03: Recent yields...")
chart03 = yields[yields["Date"] >= recent_start][["Date", "DE_sm21", "IT_sm21", "US_sm21", "UK_sm21", "JP_sm21"]].copy()
chart03.to_csv(data_dir / "chart03_recent_yields_DE_IT_outside_sm21.csv", index=False)

def plot_chart03(ax):
    cols_to_plot = ["DE_sm21", "IT_sm21", "US_sm21", "UK_sm21", "JP_sm21"]
    labels = ["DE", "IT", "US", "UK", "JP"]

    y_lim = calc_ylim([chart03[c] for c in cols_to_plot], pad=0.22)

    for col, label in zip(cols_to_plot, labels):
        ax.plot(chart03["Date"], chart03[col], linewidth=2.2,
                color=country_cols[label], label=label)

    ax.set_ylim(y_lim)
    ax.set_xlabel("Date")
    ax.set_ylabel("10 year yield (percent)")
    ax.set_title("Recent yields: Germany and Italy versus US, UK, Japan (smoothed)",
                 fontsize=11, fontweight="bold")
    ax.grid(True, color="#E6E6E6", linestyle="-", linewidth=0.5)
    ax.legend(loc="upper right", frameon=False)

    ev = filter_events(chart03["Date"].min(), chart03["Date"].max(),
                       include_ids=[2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
    add_events(ax, ev, y_lim)

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_major_locator(mdates.YearLocator(2))

ev03 = filter_events(chart03["Date"].min(), chart03["Date"].max(),
                     include_ids=[2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
plot_with_footer(
    out_dir / "chart03_recent_yields_DE_IT_outside_sm21.png",
    plot_chart03,
    [
        "What it shows: post-2014 rate dynamics with DE representing EA core and IT representing EA periphery.",
        f"Method: daily 10 year yields smoothed with a {k_yields_short}-day centred moving average.",
        event_key_text(ev03)
    ]
)

# ============================================================
# CHART 04: Fragmentation vs dispersion (standardised)
# ============================================================
print("Generating Chart 04: Fragmentation vs dispersion...")
df4 = yields[yields["Date"] >= recent_start][["Date", "IT_DE_spread_bp_sm21", "EA_dispersion_sd_sm21"]].copy()
df4["IT_DE_z"] = zscore(df4["IT_DE_spread_bp_sm21"])
df4["EA_disp_z"] = zscore(df4["EA_dispersion_sd_sm21"])
chart04 = df4[["Date", "IT_DE_z", "EA_disp_z"]].copy()
chart04.to_csv(data_dir / "chart04_fragmentation_vs_dispersion_z.csv", index=False)

def plot_chart04(ax):
    y_lim = calc_ylim_sym([chart04["IT_DE_z"], chart04["EA_disp_z"]], pad=0.25)

    ax.plot(chart04["Date"], chart04["IT_DE_z"], linewidth=2.2,
            color=country_cols["IT"], label="IT-DE spread (z)")
    ax.plot(chart04["Date"], chart04["EA_disp_z"], linewidth=2.2,
            color="#444444", label="EA dispersion (z)")
    ax.axhline(0, linestyle=":", color="#888888", linewidth=1)

    ax.set_ylim(y_lim)
    ax.set_xlabel("Date")
    ax.set_ylabel("Standardised index")
    ax.set_title("EA fragmentation and cross-country dispersion (standardised)",
                 fontsize=11, fontweight="bold")
    ax.grid(True, color="#E6E6E6", linestyle="-", linewidth=0.5)
    ax.legend(loc="upper right", frameon=False)

    ev = filter_events(chart04["Date"].min(), chart04["Date"].max(),
                       include_ids=[2, 5, 6, 7, 8, 10, 11, 14])
    add_events(ax, ev, y_lim)

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_major_locator(mdates.YearLocator(2))

ev04 = filter_events(chart04["Date"].min(), chart04["Date"].max(),
                     include_ids=[2, 5, 6, 7, 8, 10, 11, 14])
plot_with_footer(
    out_dir / "chart04_fragmentation_vs_dispersion_z.png",
    plot_chart04,
    [
        "What it shows: whether EA dispersion co-moves with fragmentation in Italy spreads.",
        f"Method: IT-DE (bp) and EA cross-sectional SD are smoothed ({k_yields_short} days) then standardised.",
        event_key_text(ev04)
    ]
)

print("\n" + "="*60)
print("COMPLETE!")
print("="*60)
print(f"\nPython charts saved to: {out_dir}")
print(f"Chart data saved to: {data_dir}")
print("\nCharts generated:")
print("  - chart01_longrun_yields_sm63.png")
print("  - chart02_EA_spreads_to_DE_bp_sm63.png")
print("  - chart03_recent_yields_DE_IT_outside_sm21.png")
print("  - chart04_fragmentation_vs_dispersion_z.png")
print("\nNote: Charts 05-13 require the external rolling spillover metrics file")
print("      which is not available in this repository.")
