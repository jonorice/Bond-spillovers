"""
Comprehensive Euro Area Bond Yield Spillover Analysis
======================================================

This script implements:
1. Generalized FEVD spillover methodology (Diebold-Yilmaz 2012, order-invariant)
2. Rolling window estimation for time-varying spillovers
3. Additional data fetching (VIX, FX rates, policy rates)
4. Network analysis and visualization
5. Regime analysis and interpretation

Author: Claude
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import FancyArrowPatch
from pathlib import Path
from datetime import datetime, timedelta
import warnings
import requests
import json
from io import StringIO

warnings.filterwarnings('ignore')

# statsmodels for VAR estimation
from statsmodels.tsa.api import VAR
from scipy import stats
import networkx as nx

# =============================================================================
# CONFIGURATION
# =============================================================================
BASE_DIR = Path("/home/user/Bond-spillovers")
OUT_DIR = BASE_DIR / "comprehensive_analysis"
OUT_DIR.mkdir(exist_ok=True)

# VAR parameters
VAR_LAGS = 4  # Common choice for daily data
FORECAST_HORIZON = 10  # FEVD horizon
ROLLING_WINDOW = 200  # Trading days (~10 months)

# Country groups
EXTERNAL = ["US", "JP", "UK"]
EA_CORE = ["DE", "FR"]
EA_PERIPHERY = ["IT", "ES"]
EA = EA_CORE + EA_PERIPHERY
ALL_COUNTRIES = EXTERNAL + EA

# Color scheme
COLORS = {
    "US": "#003299", "JP": "#FFB400", "UK": "#FF4B00",
    "DE": "#65B800", "FR": "#00B1EA", "IT": "#007816", "ES": "#8139C6"
}

# =============================================================================
# DATA FETCHING FUNCTIONS
# =============================================================================

def fetch_fred_series(series_id, start_date="1990-01-01"):
    """Fetch data from FRED (Federal Reserve Economic Data)"""
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}&cosd={start_date}"
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            df = pd.read_csv(StringIO(response.text))
            # Handle different column name formats
            df.columns = ['Date', series_id]
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.replace('.', np.nan)
            df[series_id] = pd.to_numeric(df[series_id], errors='coerce')
            df = df.dropna(subset=['Date'])
            return df
        else:
            print(f"  Failed to fetch {series_id}: HTTP {response.status_code}")
            return None
    except Exception as e:
        print(f"  Error fetching {series_id}: {e}")
        return None

def fetch_additional_data():
    """Fetch VIX, exchange rates, and policy rates from FRED"""
    print("\nFetching additional data from FRED...")

    series = {
        'VIXCLS': 'VIX',           # VIX volatility index
        'DEXUSEU': 'EURUSD',       # EUR/USD exchange rate
        'DEXUSUK': 'GBPUSD',       # GBP/USD exchange rate
        'DEXJPUS': 'USDJPY',       # USD/JPY exchange rate
        'DFF': 'FedFunds',         # Federal Funds Rate
        'ECBMRRFR': 'ECB_MRO',     # ECB Main Refinancing Rate
        'IRSTCB01JPM156N': 'BoJ_Rate',  # BoJ Policy Rate
        'BOGZ1FL073161113Q': 'ECB_Assets',  # ECB Balance Sheet (quarterly)
    }

    all_data = []
    for fred_id, name in series.items():
        print(f"  Fetching {name} ({fred_id})...")
        df = fetch_fred_series(fred_id)
        if df is not None:
            df.columns = ['Date', name]
            all_data.append(df)

    if not all_data:
        print("  Warning: Could not fetch any additional data")
        return None

    # Merge all series
    result = all_data[0]
    for df in all_data[1:]:
        result = pd.merge(result, df, on='Date', how='outer')

    result = result.sort_values('Date').reset_index(drop=True)
    print(f"  Fetched {len(result)} observations with {len(result.columns)-1} series")
    return result

# =============================================================================
# GENERALIZED FEVD SPILLOVER METHODOLOGY
# =============================================================================

def generalized_fevd(var_result, H=10):
    """
    Compute Generalized Forecast Error Variance Decomposition
    Based on Pesaran & Shin (1998), used in Diebold-Yilmaz (2012)

    This is ORDER-INVARIANT (doesn't depend on variable ordering)

    Parameters:
    -----------
    var_result : VAR results object from statsmodels
    H : int, forecast horizon

    Returns:
    --------
    theta : ndarray, normalized GFEVD matrix (rows sum to 100)
    """
    # Get VAR parameters
    k = var_result.neqs  # Number of variables
    p = var_result.k_ar  # Number of lags

    # Coefficient matrices
    A = var_result.coefs  # Shape: (p, k, k)

    # Residual covariance matrix (convert to numpy if DataFrame)
    sigma = var_result.sigma_u
    if hasattr(sigma, 'values'):
        sigma = sigma.values
    sigma_diag = np.diag(sigma)

    # Compute MA representation coefficients (Phi matrices)
    # Phi_0 = I, Phi_s = sum_{j=1}^{min(s,p)} Phi_{s-j} @ A_j
    Phi = np.zeros((H + 1, k, k))
    Phi[0] = np.eye(k)

    for s in range(1, H + 1):
        for j in range(1, min(s, p) + 1):
            Phi[s] += Phi[s - j] @ A[j - 1]

    # Compute generalized FEVD
    # theta_ij^g(H) = (sigma_jj^{-1} * sum_{h=0}^{H-1} (e_i' Phi_h Sigma e_j)^2) /
    #                 (sum_{h=0}^{H-1} e_i' Phi_h Sigma Phi_h' e_i)

    theta = np.zeros((k, k))

    for i in range(k):
        # Denominator: total forecast error variance for variable i
        denom = 0
        for h in range(H):
            denom += Phi[h][i, :] @ sigma @ Phi[h][i, :].T

        for j in range(k):
            # Numerator: contribution from shock j to variable i
            numer = 0
            for h in range(H):
                numer += (Phi[h][i, :] @ sigma[:, j]) ** 2

            theta[i, j] = (numer / sigma_diag[j]) / denom if denom > 0 else 0

    # Normalize rows to sum to 100 (as in Diebold-Yilmaz)
    row_sums = theta.sum(axis=1, keepdims=True)
    row_sums[row_sums == 0] = 1  # Avoid division by zero
    theta_normalized = 100 * theta / row_sums

    return theta_normalized


def compute_spillover_indices(theta, var_names):
    """
    Compute various spillover indices from the GFEVD matrix

    Parameters:
    -----------
    theta : ndarray, GFEVD matrix (rows sum to 100)
    var_names : list, variable names

    Returns:
    --------
    dict with spillover measures
    """
    k = len(var_names)

    # Total Spillover Index
    # = (sum of off-diagonal elements) / k * 100
    off_diag_sum = theta.sum() - np.trace(theta)
    total_spillover = off_diag_sum / k

    # Directional spillovers TO others (column sums minus diagonal)
    to_others = theta.sum(axis=0) - np.diag(theta)

    # Directional spillovers FROM others (row sums minus diagonal)
    from_others = theta.sum(axis=1) - np.diag(theta)

    # Net spillovers (to - from)
    net = to_others - from_others

    # Pairwise spillovers
    pairwise = {}
    for i, name_i in enumerate(var_names):
        for j, name_j in enumerate(var_names):
            if i != j:
                pairwise[f"{name_j}_to_{name_i}"] = theta[i, j]

    return {
        'total_spillover': total_spillover,
        'to_others': dict(zip(var_names, to_others)),
        'from_others': dict(zip(var_names, from_others)),
        'net': dict(zip(var_names, net)),
        'pairwise': pairwise,
        'theta': theta
    }


def rolling_spillover_analysis(data, var_names, window=200, lag=4, horizon=10, step=5):
    """
    Compute rolling window spillover analysis

    Parameters:
    -----------
    data : DataFrame with yields
    var_names : list of column names to include
    window : rolling window size
    lag : VAR lag order
    horizon : FEVD horizon
    step : step size for rolling (to speed up computation)

    Returns:
    --------
    DataFrame with rolling spillover measures
    """
    results = []
    n = len(data)

    print(f"\nComputing rolling spillovers (window={window}, step={step})...")
    total_windows = (n - window) // step + 1

    for i, start in enumerate(range(0, n - window, step)):
        if i % 50 == 0:
            print(f"  Processing window {i+1}/{total_windows} ({100*i/total_windows:.0f}%)")

        end = start + window
        window_data = data.iloc[start:end][var_names].dropna()

        if len(window_data) < window * 0.9:  # Require at least 90% non-missing
            continue

        try:
            # Estimate VAR
            model = VAR(window_data)
            var_result = model.fit(lag)

            # Compute GFEVD
            theta = generalized_fevd(var_result, H=horizon)

            # Compute indices
            indices = compute_spillover_indices(theta, var_names)

            # Store results
            result = {
                'Date': data.iloc[end - 1]['Date'],
                'total_spillover': indices['total_spillover']
            }

            # Add directional measures
            for name in var_names:
                result[f'{name}_to'] = indices['to_others'][name]
                result[f'{name}_from'] = indices['from_others'][name]
                result[f'{name}_net'] = indices['net'][name]

            # Add key pairwise measures
            for key, val in indices['pairwise'].items():
                result[key] = val

            results.append(result)

        except Exception as e:
            continue

    print(f"  Completed {len(results)} windows")
    return pd.DataFrame(results)


# =============================================================================
# ANALYSIS FUNCTIONS
# =============================================================================

def compute_regime_statistics(roll_df, regimes):
    """Compute average spillovers by regime"""
    stats = []
    for regime_name, (start, end) in regimes.items():
        mask = (roll_df['Date'] >= start) & (roll_df['Date'] <= end)
        regime_data = roll_df[mask]

        if len(regime_data) == 0:
            continue

        stat = {'regime': regime_name, 'n_obs': len(regime_data)}

        # Total spillover
        stat['total_spillover'] = regime_data['total_spillover'].mean()

        # Net spillovers by country
        for c in ALL_COUNTRIES:
            if f'{c}_net' in regime_data.columns:
                stat[f'{c}_net'] = regime_data[f'{c}_net'].mean()

        stats.append(stat)

    return pd.DataFrame(stats)


def compute_external_vs_internal(roll_df):
    """Compute spillovers from external vs within-EA sources"""
    # Spillovers INTO EA countries FROM external sources
    external_to_ea_cols = []
    internal_ea_cols = []

    for ea_c in EA:
        for ext_c in EXTERNAL:
            col = f'{ext_c}_to_{ea_c}'
            if col in roll_df.columns:
                external_to_ea_cols.append(col)

        for other_ea in EA:
            if other_ea != ea_c:
                col = f'{other_ea}_to_{ea_c}'
                if col in roll_df.columns:
                    internal_ea_cols.append(col)

    result = roll_df[['Date']].copy()

    if external_to_ea_cols:
        result['external_to_EA'] = roll_df[external_to_ea_cols].mean(axis=1)

    if internal_ea_cols:
        result['within_EA'] = roll_df[internal_ea_cols].mean(axis=1)

    # US specifically to EA
    us_to_ea_cols = [f'US_to_{ea_c}' for ea_c in EA if f'US_to_{ea_c}' in roll_df.columns]
    if us_to_ea_cols:
        result['US_to_EA'] = roll_df[us_to_ea_cols].mean(axis=1)

    # JP to EA
    jp_to_ea_cols = [f'JP_to_{ea_c}' for ea_c in EA if f'JP_to_{ea_c}' in roll_df.columns]
    if jp_to_ea_cols:
        result['JP_to_EA'] = roll_df[jp_to_ea_cols].mean(axis=1)

    # UK to EA
    uk_to_ea_cols = [f'UK_to_{ea_c}' for ea_c in EA if f'UK_to_{ea_c}' in roll_df.columns]
    if uk_to_ea_cols:
        result['UK_to_EA'] = roll_df[uk_to_ea_cols].mean(axis=1)

    return result


def compute_fragmentation_index(roll_df):
    """Compute EA fragmentation based on spillover asymmetries"""
    result = roll_df[['Date']].copy()

    # Core to periphery vs periphery to core
    core_to_periph = []
    periph_to_core = []

    for core in EA_CORE:
        for periph in EA_PERIPHERY:
            c2p = f'{core}_to_{periph}'
            p2c = f'{periph}_to_{core}'
            if c2p in roll_df.columns:
                core_to_periph.append(c2p)
            if p2c in roll_df.columns:
                periph_to_core.append(p2c)

    if core_to_periph:
        result['core_to_periphery'] = roll_df[core_to_periph].mean(axis=1)
    if periph_to_core:
        result['periphery_to_core'] = roll_df[periph_to_core].mean(axis=1)

    # Net position of periphery (negative = receiving more)
    if 'core_to_periphery' in result.columns and 'periphery_to_core' in result.columns:
        result['periphery_net'] = result['periphery_to_core'] - result['core_to_periphery']

    return result


# =============================================================================
# VISUALIZATION FUNCTIONS
# =============================================================================

def plot_total_spillover(roll_df, events, out_path):
    """Plot total spillover index over time"""
    fig, ax = plt.subplots(figsize=(14, 6))

    ax.plot(roll_df['Date'], roll_df['total_spillover'],
            color='#003299', linewidth=1.5, label='Total Spillover Index')

    ax.fill_between(roll_df['Date'], roll_df['total_spillover'],
                    alpha=0.3, color='#003299')

    # Add events
    for _, ev in events.iterrows():
        if ev['date'] >= roll_df['Date'].min() and ev['date'] <= roll_df['Date'].max():
            ax.axvline(ev['date'], color='gray', linestyle=':', alpha=0.5)
            ax.text(ev['date'], ax.get_ylim()[1], str(ev['id']),
                    fontsize=8, ha='center', va='bottom')

    ax.set_xlabel('Date')
    ax.set_ylabel('Total Spillover Index (%)')
    ax.set_title('Diebold-Yilmaz Total Spillover Index (Generalized FEVD)', fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.legend(loc='upper left')

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_net_spillovers(roll_df, out_path):
    """Plot net spillovers by country"""
    fig, axes = plt.subplots(2, 1, figsize=(14, 10), sharex=True)

    # External countries
    ax1 = axes[0]
    for c in EXTERNAL:
        col = f'{c}_net'
        if col in roll_df.columns:
            ax1.plot(roll_df['Date'], roll_df[col],
                    color=COLORS[c], linewidth=1.5, label=c)
    ax1.axhline(0, color='black', linestyle='-', linewidth=0.5)
    ax1.set_ylabel('Net Spillover')
    ax1.set_title('Net Spillovers: External Countries (+ = net transmitter)', fontweight='bold')
    ax1.legend(loc='upper right')
    ax1.grid(True, alpha=0.3)

    # EA countries
    ax2 = axes[1]
    for c in EA:
        col = f'{c}_net'
        if col in roll_df.columns:
            ax2.plot(roll_df['Date'], roll_df[col],
                    color=COLORS[c], linewidth=1.5, label=c)
    ax2.axhline(0, color='black', linestyle='-', linewidth=0.5)
    ax2.set_xlabel('Date')
    ax2.set_ylabel('Net Spillover')
    ax2.set_title('Net Spillovers: Euro Area Countries (+ = net transmitter)', fontweight='bold')
    ax2.legend(loc='upper right')
    ax2.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_external_vs_internal(ext_int_df, out_path):
    """Plot external vs internal EA spillovers"""
    fig, ax = plt.subplots(figsize=(14, 6))

    if 'external_to_EA' in ext_int_df.columns:
        ax.plot(ext_int_df['Date'], ext_int_df['external_to_EA'],
                color='#222222', linewidth=2, label='From External (US+JP+UK)')

    if 'within_EA' in ext_int_df.columns:
        ax.plot(ext_int_df['Date'], ext_int_df['within_EA'],
                color='#777777', linewidth=2, label='Within EA')

    ax.set_xlabel('Date')
    ax.set_ylabel('Average Spillover to EA Members (%)')
    ax.set_title('Spillovers into Euro Area: External vs Internal Sources', fontweight='bold')
    ax.legend(loc='upper right')
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_external_breakdown(ext_int_df, out_path):
    """Plot external spillovers by source country"""
    fig, ax = plt.subplots(figsize=(14, 6))

    for c in EXTERNAL:
        col = f'{c}_to_EA'
        if col in ext_int_df.columns:
            ax.plot(ext_int_df['Date'], ext_int_df[col],
                    color=COLORS[c], linewidth=1.5, label=f'{c} → EA')

    ax.set_xlabel('Date')
    ax.set_ylabel('Average Spillover to EA Members (%)')
    ax.set_title('External Spillovers into Euro Area by Source', fontweight='bold')
    ax.legend(loc='upper right')
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_fragmentation(frag_df, out_path):
    """Plot core-periphery spillover dynamics"""
    fig, axes = plt.subplots(2, 1, figsize=(14, 10), sharex=True)

    ax1 = axes[0]
    if 'core_to_periphery' in frag_df.columns:
        ax1.plot(frag_df['Date'], frag_df['core_to_periphery'],
                color=COLORS['DE'], linewidth=1.5, label='Core → Periphery')
    if 'periphery_to_core' in frag_df.columns:
        ax1.plot(frag_df['Date'], frag_df['periphery_to_core'],
                color=COLORS['IT'], linewidth=1.5, label='Periphery → Core')
    ax1.set_ylabel('Average Spillover (%)')
    ax1.set_title('Core-Periphery Spillover Flows', fontweight='bold')
    ax1.legend(loc='upper right')
    ax1.grid(True, alpha=0.3)

    ax2 = axes[1]
    if 'periphery_net' in frag_df.columns:
        ax2.fill_between(frag_df['Date'], frag_df['periphery_net'],
                        where=frag_df['periphery_net'] >= 0, color=COLORS['IT'], alpha=0.5)
        ax2.fill_between(frag_df['Date'], frag_df['periphery_net'],
                        where=frag_df['periphery_net'] < 0, color=COLORS['DE'], alpha=0.5)
        ax2.plot(frag_df['Date'], frag_df['periphery_net'], color='black', linewidth=1)
    ax2.axhline(0, color='black', linestyle='-', linewidth=0.5)
    ax2.set_xlabel('Date')
    ax2.set_ylabel('Net Spillover')
    ax2.set_title('Periphery Net Position (+ = net transmitter to core)', fontweight='bold')
    ax2.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_spillover_network(theta, var_names, title, out_path, threshold=5):
    """Plot network graph of spillovers"""
    fig, ax = plt.subplots(figsize=(10, 10))

    G = nx.DiGraph()

    # Add nodes
    for name in var_names:
        G.add_node(name)

    # Add edges (spillovers above threshold)
    for i, from_name in enumerate(var_names):
        for j, to_name in enumerate(var_names):
            if i != j and theta[j, i] > threshold:  # theta[j,i] = spillover from i to j
                G.add_edge(from_name, to_name, weight=theta[j, i])

    # Position nodes in a circle
    pos = nx.circular_layout(G)

    # Draw nodes
    node_colors = [COLORS.get(n, '#888888') for n in G.nodes()]
    nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=2000, ax=ax)
    nx.draw_networkx_labels(G, pos, font_size=12, font_weight='bold', ax=ax)

    # Draw edges with varying width
    edges = G.edges(data=True)
    for (u, v, d) in edges:
        width = d['weight'] / 5
        alpha = min(d['weight'] / 20, 0.8)
        nx.draw_networkx_edges(G, pos, edgelist=[(u, v)], width=width,
                              alpha=alpha, edge_color='gray',
                              connectionstyle='arc3,rad=0.1', ax=ax,
                              arrows=True, arrowsize=20)

    ax.set_title(title, fontsize=14, fontweight='bold')
    ax.axis('off')

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_regime_comparison(regime_stats, out_path):
    """Plot regime comparison bar chart"""
    if len(regime_stats) == 0:
        print("No regime statistics to plot")
        return

    fig, axes = plt.subplots(2, 1, figsize=(14, 10))

    # Total spillover by regime
    ax1 = axes[0]
    regimes = regime_stats['regime'].tolist()
    totals = regime_stats['total_spillover'].tolist()
    bars = ax1.bar(regimes, totals, color='#003299', alpha=0.7)
    ax1.set_ylabel('Total Spillover Index (%)')
    ax1.set_title('Total Spillover by Regime', fontweight='bold')
    ax1.grid(True, alpha=0.3, axis='y')

    # Add value labels
    for bar, val in zip(bars, totals):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                f'{val:.1f}', ha='center', va='bottom', fontsize=10)

    # Net spillovers by country and regime
    ax2 = axes[1]
    x = np.arange(len(regimes))
    width = 0.12

    for i, c in enumerate(ALL_COUNTRIES):
        col = f'{c}_net'
        if col in regime_stats.columns:
            vals = regime_stats[col].tolist()
            ax2.bar(x + i*width, vals, width, label=c, color=COLORS[c], alpha=0.8)

    ax2.axhline(0, color='black', linewidth=0.5)
    ax2.set_xlabel('Regime')
    ax2.set_ylabel('Net Spillover')
    ax2.set_title('Net Spillovers by Country and Regime', fontweight='bold')
    ax2.set_xticks(x + width * 3)
    ax2.set_xticklabels(regimes)
    ax2.legend(loc='upper right', ncol=4)
    ax2.grid(True, alpha=0.3, axis='y')

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


def plot_vix_spillover_relationship(roll_df, vix_data, out_path):
    """Plot relationship between VIX and total spillovers"""
    # Merge data
    merged = pd.merge(roll_df[['Date', 'total_spillover']],
                      vix_data[['Date', 'VIX']], on='Date', how='inner')
    merged = merged.dropna()

    if len(merged) < 50:
        print("Insufficient overlapping data for VIX analysis")
        return

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))

    # Time series
    ax1 = axes[0, 0]
    ax1b = ax1.twinx()
    l1, = ax1.plot(merged['Date'], merged['total_spillover'], color='#003299',
                   linewidth=1.5, label='Spillover Index')
    l2, = ax1b.plot(merged['Date'], merged['VIX'], color='#FF4B00',
                    linewidth=1.5, alpha=0.7, label='VIX')
    ax1.set_ylabel('Total Spillover (%)', color='#003299')
    ax1b.set_ylabel('VIX', color='#FF4B00')
    ax1.set_title('Spillovers and VIX Over Time', fontweight='bold')
    ax1.legend(handles=[l1, l2], loc='upper right')

    # Scatter plot
    ax2 = axes[0, 1]
    ax2.scatter(merged['VIX'], merged['total_spillover'], alpha=0.3, s=10)
    # Add regression line
    z = np.polyfit(merged['VIX'], merged['total_spillover'], 1)
    p = np.poly1d(z)
    vix_range = np.linspace(merged['VIX'].min(), merged['VIX'].max(), 100)
    ax2.plot(vix_range, p(vix_range), 'r-', linewidth=2, label=f'Slope: {z[0]:.2f}')
    corr = merged['VIX'].corr(merged['total_spillover'])
    ax2.set_xlabel('VIX')
    ax2.set_ylabel('Total Spillover (%)')
    ax2.set_title(f'Spillover vs VIX (Corr: {corr:.3f})', fontweight='bold')
    ax2.legend()
    ax2.grid(True, alpha=0.3)

    # Rolling correlation
    ax3 = axes[1, 0]
    merged['rolling_corr'] = merged['total_spillover'].rolling(50).corr(merged['VIX'])
    ax3.plot(merged['Date'], merged['rolling_corr'], color='purple', linewidth=1.5)
    ax3.axhline(0, color='black', linestyle='-', linewidth=0.5)
    ax3.set_xlabel('Date')
    ax3.set_ylabel('Rolling Correlation')
    ax3.set_title('Rolling Correlation: Spillovers vs VIX (50-day)', fontweight='bold')
    ax3.grid(True, alpha=0.3)

    # VIX regime comparison
    ax4 = axes[1, 1]
    merged['vix_regime'] = pd.cut(merged['VIX'], bins=[0, 15, 25, 100],
                                   labels=['Low (<15)', 'Medium (15-25)', 'High (>25)'])
    regime_means = merged.groupby('vix_regime')['total_spillover'].mean()
    bars = ax4.bar(regime_means.index.astype(str), regime_means.values,
                   color=['green', 'orange', 'red'], alpha=0.7)
    ax4.set_xlabel('VIX Regime')
    ax4.set_ylabel('Average Total Spillover (%)')
    ax4.set_title('Spillovers by VIX Regime', fontweight='bold')
    ax4.grid(True, alpha=0.3, axis='y')

    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Saved: {out_path}")


# =============================================================================
# MAIN ANALYSIS
# =============================================================================

def main():
    print("="*70)
    print("COMPREHENSIVE EURO AREA BOND YIELD SPILLOVER ANALYSIS")
    print("="*70)

    # -------------------------------------------------------------------------
    # Load yield data
    # -------------------------------------------------------------------------
    print("\n1. LOADING DATA")
    print("-"*50)

    yields = pd.read_csv(BASE_DIR / "nav_test_clean.csv")
    yields['Date'] = pd.to_datetime(yields['Date'])
    yields = yields.sort_values('Date').reset_index(drop=True)

    # Convert to numeric
    for c in ALL_COUNTRIES:
        yields[c] = pd.to_numeric(yields[c], errors='coerce')

    print(f"Loaded yields: {len(yields)} observations")
    print(f"Date range: {yields['Date'].min().date()} to {yields['Date'].max().date()}")

    # Fetch additional data
    additional_data = fetch_additional_data()

    # -------------------------------------------------------------------------
    # Define events
    # -------------------------------------------------------------------------
    events = pd.DataFrame({
        "id": list(range(1, 15)),
        "date": pd.to_datetime([
            "2012-07-26", "2015-01-22", "2016-06-23", "2016-09-21",
            "2018-05-29", "2019-09-12", "2020-03-18", "2022-02-24",
            "2022-03-16", "2022-07-21", "2022-09-23", "2022-12-20",
            "2024-03-19", "2024-06-06"
        ]),
        "label": [
            "Draghi 'whatever it takes'", "ECB QE announced", "Brexit referendum",
            "BoJ YCC introduced", "Italy BTP crisis", "ECB restarts QE",
            "ECB PEPP", "Ukraine invasion", "Fed hiking begins",
            "ECB lift-off + TPI", "UK mini-budget", "BoJ widens YCC",
            "BoJ ends NIRP/YCC", "ECB cutting cycle"
        ]
    })
    events.to_csv(OUT_DIR / "events.csv", index=False)

    # Define regimes
    regimes = {
        'Pre-Draghi (2010-2012)': (pd.Timestamp('2010-01-01'), pd.Timestamp('2012-07-25')),
        'QE Era (2015-2019)': (pd.Timestamp('2015-01-01'), pd.Timestamp('2019-12-31')),
        'COVID/PEPP (2020-2021)': (pd.Timestamp('2020-01-01'), pd.Timestamp('2021-12-31')),
        'Tightening (2022-2023)': (pd.Timestamp('2022-01-01'), pd.Timestamp('2023-12-31')),
        'Cutting Cycle (2024+)': (pd.Timestamp('2024-01-01'), pd.Timestamp('2025-12-31'))
    }

    # -------------------------------------------------------------------------
    # Compute rolling spillovers
    # -------------------------------------------------------------------------
    print("\n2. COMPUTING SPILLOVER ANALYSIS")
    print("-"*50)

    # Use yield changes (first differences) - more stationary
    yield_changes = yields.copy()
    for c in ALL_COUNTRIES:
        yield_changes[c] = yields[c].diff()
    yield_changes = yield_changes.dropna().reset_index(drop=True)

    print(f"Using yield changes: {len(yield_changes)} observations")

    # Run rolling spillover analysis
    roll_df = rolling_spillover_analysis(
        yield_changes, ALL_COUNTRIES,
        window=ROLLING_WINDOW, lag=VAR_LAGS, horizon=FORECAST_HORIZON, step=5
    )

    roll_df.to_csv(OUT_DIR / "rolling_spillovers.csv", index=False)
    print(f"Rolling spillovers saved: {len(roll_df)} windows")

    # -------------------------------------------------------------------------
    # Compute derived measures
    # -------------------------------------------------------------------------
    print("\n3. COMPUTING DERIVED MEASURES")
    print("-"*50)

    ext_int_df = compute_external_vs_internal(roll_df)
    ext_int_df.to_csv(OUT_DIR / "external_vs_internal.csv", index=False)
    print("External vs internal spillovers computed")

    frag_df = compute_fragmentation_index(roll_df)
    frag_df.to_csv(OUT_DIR / "fragmentation.csv", index=False)
    print("Fragmentation measures computed")

    regime_stats = compute_regime_statistics(roll_df, regimes)
    regime_stats.to_csv(OUT_DIR / "regime_statistics.csv", index=False)
    print("Regime statistics computed")

    # -------------------------------------------------------------------------
    # Generate visualizations
    # -------------------------------------------------------------------------
    print("\n4. GENERATING VISUALIZATIONS")
    print("-"*50)

    plot_total_spillover(roll_df, events, OUT_DIR / "fig01_total_spillover.png")
    plot_net_spillovers(roll_df, OUT_DIR / "fig02_net_spillovers.png")
    plot_external_vs_internal(ext_int_df, OUT_DIR / "fig03_external_vs_internal.png")
    plot_external_breakdown(ext_int_df, OUT_DIR / "fig04_external_breakdown.png")
    plot_fragmentation(frag_df, OUT_DIR / "fig05_fragmentation.png")
    plot_regime_comparison(regime_stats, OUT_DIR / "fig06_regime_comparison.png")

    # Network plots for specific periods
    print("\nComputing period-specific networks...")

    for regime_name, (start, end) in regimes.items():
        mask = (yield_changes['Date'] >= start) & (yield_changes['Date'] <= end)
        period_data = yield_changes[mask][ALL_COUNTRIES].dropna()

        if len(period_data) < 100:
            continue

        try:
            model = VAR(period_data)
            var_result = model.fit(VAR_LAGS)
            theta = generalized_fevd(var_result, H=FORECAST_HORIZON)

            safe_name = regime_name.replace(' ', '_').replace('(', '').replace(')', '').replace('+', '')
            plot_spillover_network(theta, ALL_COUNTRIES,
                                   f'Spillover Network: {regime_name}',
                                   OUT_DIR / f"fig_network_{safe_name}.png")
        except Exception as e:
            print(f"  Could not compute network for {regime_name}: {e}")

    # VIX relationship if data available
    if additional_data is not None and 'VIX' in additional_data.columns:
        print("\nAnalyzing VIX-spillover relationship...")
        plot_vix_spillover_relationship(roll_df, additional_data,
                                        OUT_DIR / "fig07_vix_relationship.png")

    # -------------------------------------------------------------------------
    # Generate insights
    # -------------------------------------------------------------------------
    print("\n5. GENERATING INSIGHTS")
    print("-"*50)

    insights = []

    # Insight 1: Overall spillover trends
    early = roll_df[roll_df['Date'] < '2015-01-01']['total_spillover'].mean()
    late = roll_df[roll_df['Date'] >= '2020-01-01']['total_spillover'].mean()
    insights.append(f"1. SPILLOVER INTENSITY: Average total spillover was {early:.1f}% pre-2015 "
                   f"vs {late:.1f}% post-2020 ({(late-early)/early*100:+.0f}% change)")

    # Insight 2: US dominance
    if 'US_to_EA' in ext_int_df.columns:
        us_max = ext_int_df['US_to_EA'].max()
        us_max_date = ext_int_df.loc[ext_int_df['US_to_EA'].idxmax(), 'Date']
        insights.append(f"2. US DOMINANCE: Peak US→EA spillover was {us_max:.1f}% on "
                       f"{us_max_date.strftime('%Y-%m-%d')} (Fed hiking cycle)")

    # Insight 3: External vs internal
    if 'external_to_EA' in ext_int_df.columns and 'within_EA' in ext_int_df.columns:
        ext_mean = ext_int_df['external_to_EA'].mean()
        int_mean = ext_int_df['within_EA'].mean()
        insights.append(f"3. EXTERNAL VS INTERNAL: On average, {ext_mean:.1f}% of EA yield "
                       f"variance comes from external sources vs {int_mean:.1f}% from within EA")

    # Insight 4: Net transmitters
    net_means = {}
    for c in ALL_COUNTRIES:
        if f'{c}_net' in roll_df.columns:
            net_means[c] = roll_df[f'{c}_net'].mean()

    if net_means:
        biggest_transmitter = max(net_means, key=net_means.get)
        biggest_receiver = min(net_means, key=net_means.get)
        insights.append(f"4. NET POSITIONS: {biggest_transmitter} is the largest net transmitter "
                       f"({net_means[biggest_transmitter]:.1f}), {biggest_receiver} is the largest "
                       f"net receiver ({net_means[biggest_receiver]:.1f})")

    # Insight 5: Fragmentation
    if 'periphery_net' in frag_df.columns:
        periph_pos = (frag_df['periphery_net'] > 0).mean() * 100
        insights.append(f"5. FRAGMENTATION: EA periphery (IT, ES) is a net transmitter to core "
                       f"only {periph_pos:.0f}% of the time - usually receives more spillovers")

    # Insight 6: Regime differences
    if len(regime_stats) >= 2:
        max_regime = regime_stats.loc[regime_stats['total_spillover'].idxmax(), 'regime']
        min_regime = regime_stats.loc[regime_stats['total_spillover'].idxmin(), 'regime']
        insights.append(f"6. REGIME VARIATION: Highest spillovers during '{max_regime}', "
                       f"lowest during '{min_regime}'")

    # Insight 7: Japan
    if 'JP_to_EA' in ext_int_df.columns:
        jp_2022_2023 = ext_int_df[(ext_int_df['Date'] >= '2022-01-01') &
                                  (ext_int_df['Date'] < '2024-01-01')]['JP_to_EA'].mean()
        jp_earlier = ext_int_df[ext_int_df['Date'] < '2022-01-01']['JP_to_EA'].mean()
        insights.append(f"7. JAPAN EFFECT: JP→EA spillovers averaged {jp_earlier:.1f}% pre-2022 "
                       f"vs {jp_2022_2023:.1f}% in 2022-23 (BoJ YCC adjustments)")

    # Insight 8: UK post-Brexit
    if 'UK_to_EA' in ext_int_df.columns:
        uk_pre_brexit = ext_int_df[ext_int_df['Date'] < '2016-06-23']['UK_to_EA'].mean()
        uk_post_brexit = ext_int_df[ext_int_df['Date'] >= '2020-01-01']['UK_to_EA'].mean()
        insights.append(f"8. POST-BREXIT UK: UK→EA spillovers were {uk_pre_brexit:.1f}% pre-Brexit "
                       f"vs {uk_post_brexit:.1f}% post-2020 ({(uk_post_brexit-uk_pre_brexit)/uk_pre_brexit*100:+.0f}%)")

    # Write insights
    with open(OUT_DIR / "KEY_INSIGHTS.txt", 'w') as f:
        f.write("="*70 + "\n")
        f.write("KEY INSIGHTS: EURO AREA BOND YIELD SPILLOVERS\n")
        f.write("="*70 + "\n\n")
        f.write("Methodology: Generalized FEVD (Diebold-Yilmaz 2012, order-invariant)\n")
        f.write(f"VAR specification: {VAR_LAGS} lags, {FORECAST_HORIZON}-day horizon\n")
        f.write(f"Rolling window: {ROLLING_WINDOW} trading days\n")
        f.write(f"Sample: {roll_df['Date'].min().strftime('%Y-%m-%d')} to "
               f"{roll_df['Date'].max().strftime('%Y-%m-%d')}\n\n")
        f.write("-"*70 + "\n\n")
        for insight in insights:
            f.write(insight + "\n\n")

    print("\nKEY INSIGHTS:")
    print("-"*50)
    for insight in insights:
        print(insight)
        print()

    print("\n" + "="*70)
    print("ANALYSIS COMPLETE")
    print("="*70)
    print(f"\nAll outputs saved to: {OUT_DIR}")
    print("\nFiles generated:")
    for f in sorted(OUT_DIR.glob("*")):
        print(f"  {f.name}")


if __name__ == "__main__":
    main()
