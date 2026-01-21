# Comprehensive Euro Area Bond Yield Spillover Analysis

## Executive Summary

This analysis uses the Generalized Forecast Error Variance Decomposition (GFEVD) methodology from Diebold-Yilmaz (2012) to measure yield spillovers across 7 sovereign bond markets: US, Japan, UK, Germany, France, Italy, and Spain. The methodology is **order-invariant**, avoiding the arbitrary ordering assumptions of traditional Cholesky decomposition.

**Key Finding**: Euro area yield dynamics are dominated by **internal transmission** (within-EA spillovers ~14% vs external ~8%), but the US influence peaks during Fed policy pivots. Spain is the most vulnerable market, while France has emerged as the "swing" market that bridges core and periphery.

---

## Methodology

| Parameter | Value |
|-----------|-------|
| Model | VAR(4) in yield changes |
| FEVD Horizon | 10 days |
| Rolling Window | 200 trading days |
| Sample | 1995-01-03 to 2025-12-12 |
| Identification | Generalized FEVD (Pesaran-Shin) |

---

## Key Findings

### 1. Total Spillover Dynamics

- **Average Total Spillover**: 60.5% of yield variance is explained by shocks from other countries
- **Range**: 45% (2010 Euro crisis trough) to 76% (2008 GFC peak)
- **Current Level**: 65-70% (elevated during tightening cycle)

**Structural Breaks**:
- 1998 (LTCM): +4.3 percentage points increase
- 2010-2012 (Euro crisis): Collapse to 50%
- 2021-2022: Rapid rise to new highs (69-70%)

### 2. External vs Internal Spillovers

| Period | External → EA | Within EA | Ratio |
|--------|--------------|-----------|-------|
| Pre-EMU (1995-98) | 8.8% | 13.7% | 0.67 |
| EMU Early (1999-2007) | 8.8% | 16.8% | 0.53 |
| GFC/Euro Crisis (2008-12) | 7.9% | 11.9% | 0.70 |
| QE Era (2013-19) | 7.3% | 11.0% | 0.68 |
| COVID (2020-21) | 7.5% | 14.3% | 0.53 |
| Tightening (2022-25) | 9.0% | 16.2% | 0.56 |

**Insight**: Within-EA spillovers consistently exceed external spillovers by 40-90%. The EA is more affected by itself than by the US, Japan, and UK combined.

### 3. Net Transmitters vs Receivers

| Country | Avg Net Spillover | Role |
|---------|------------------|------|
| **France** | +22.6 | Largest transmitter |
| **Germany** | +21.5 | Major transmitter |
| **US** | +11.1 | External transmitter |
| **UK** | +8.7 | External transmitter |
| **Italy** | +6.5 | Balanced |
| **Japan** | -26.4 | Net receiver |
| **Spain** | -39.0 | Largest receiver |

**Key Insight**: France has overtaken Germany as the dominant EA transmitter since 2015. Spain is extremely vulnerable - the largest net receiver in the system.

### 4. Bilateral Spillover Matrix (Full Sample)

```
Receiver ↓  |  US   JP   UK   DE   FR   IT   ES  (Transmitter →)
------------|--------------------------------------------
US          |   -  2.2  14.0 17.0 14.0  8.1  2.6
JP          | 8.7   -   6.0  7.0  6.3  4.5  3.1
UK          |13.7  2.0   -  18.1 16.4  9.8  4.1
DE          |14.0  1.7  15.8  -  21.8 12.1  4.4
FR          |11.1  1.7  14.2 22.2   - 13.8  4.6
IT          | 8.3  1.6   9.9 14.4 17.0   -  9.4
ES          | 9.6  2.0   9.3 12.9 14.7 18.8   -
```

**Insights**:
- Germany → France: 22.2% (strongest bilateral link)
- France → Germany: 21.8% (near-symmetric)
- Italy → Spain: 18.8% (periphery contagion channel)
- Japan transmits very little to anyone (1.6-2.2%)

### 5. Regime-Specific Patterns

#### Euro Crisis (2010-2012)
- Total spillover collapsed to 54%
- Periphery contagion spiked: IT→ES = 25.4%, ES→IT = 17.4%
- Germany-France link remained strong (21-22%)
- External influence (US) remained stable at ~16%

#### QE Era (2015-2019)
- Gradual re-integration (spillovers rose to 54%)
- Germany influence recovered
- Italy-Spain link weakened (IT→ES fell to 17.9%)

#### Tightening Cycle (2022-2025)
- Highest spillovers on record (68-70%)
- Very uniform bilateral links (~15-23% across all pairs)
- US influence at 30-year high (15% to EA)
- VIX correlation turned positive (r=0.14)

### 6. Core-Periphery Dynamics

| Period | Core → Periphery | Periphery → Core | Net |
|--------|-----------------|------------------|-----|
| Pre-EMU | 13.4% | 6.5% | -6.9 |
| EMU Early | 19.6% | 13.3% | -6.3 |
| GFC | 16.7% | 10.7% | -6.0 |
| Euro Crisis | 5.8% | 3.5% | -2.4 |
| QE Era | 10.1% | 5.1% | -5.0 |
| COVID | 17.0% | 7.5% | -9.5 |
| Tightening | 20.3% | 11.3% | -9.0 |

**Insight**: The periphery is ALWAYS a net receiver from the core. The gap narrowed only during the euro crisis when all links weakened. It has widened again post-2020.

### 7. Japan's Unique Role

- Japan is the most isolated market globally
- YCC (2016-2022) actually INCREASED Japan's spillover engagement
- Post-NIRP exit (2024): Japan became even MORE disconnected
- JP→EA: only 1-2% vs EA→JP: 5-7%

### 8. VIX-Spillover Relationship

- Overall correlation: +0.14 (weak positive)
- Rolling correlation ranges from -0.6 to +0.8
- High VIX regimes (>25) have 3% higher spillovers than low VIX (<15)
- Relationship is regime-dependent, not stable

---

## Narrative Themes for Your Paper

### Theme 1: "The Price-Taker Periphery"
Italy and Spain are structurally positioned as receivers in the EA system. Their yield dynamics are largely determined by shocks from Germany and France, not the reverse. Policy implication: Peripheral fiscal credibility matters less for EA-wide dynamics than core country shocks.

### Theme 2: "France as the Hidden Anchor"
France has become the "swing" market in EA yield dynamics, transmitting more than Germany since 2015. This may reflect France's position bridging core and periphery, or the depth/liquidity of the OAT market.

### Theme 3: "The US Shadow"
US influence peaks during Fed policy shifts. During the 2022-23 hiking cycle, US spillovers to Germany reached 15% - the highest since 1998. The EA is not isolated from US monetary policy.

### Theme 4: "Integration Paradox"
Total spillovers are LOWEST during crises (euro crisis: 50%) and HIGHEST during tightening cycles (68-70%). Markets fragment when stress is idiosyncratic (sovereign risk) but integrate when facing common shocks (global rates).

### Theme 5: "The Japan Anomaly"
Japan remains largely disconnected from global bond markets despite decades of financial integration. BoJ policy regimes matter: YCC increased linkages, while the 2024 policy exit reduced them.

---

## Files Generated

| File | Description |
|------|-------------|
| `rolling_spillovers.csv` | Daily rolling spillover measures |
| `external_vs_internal.csv` | External vs within-EA decomposition |
| `fragmentation.csv` | Core-periphery dynamics |
| `regime_statistics.csv` | Regime averages |
| `fig01-09` | Visualizations |
| `KEY_INSIGHTS.txt` | Summary statistics |

---

## Technical Notes

1. **Order Invariance**: The Generalized FEVD uses the historical error correlation structure rather than imposing Cholesky orthogonalization. Results are robust to variable ordering.

2. **Stationarity**: Analysis uses yield changes (first differences) rather than levels to ensure stationarity.

3. **Rolling Windows**: 200-day windows (~10 months) balance stability with responsiveness. Results are smoothed with 20-day moving averages for visualization.

4. **Interpretation**: A spillover of X% means X% of the h-step forecast error variance of the receiver is attributable to shocks from the transmitter.

---

*Analysis conducted using Python with statsmodels VAR implementation and Generalized FEVD.*
