"""
TCS Financial Model — Python Automation Layer
==============================================
Project : TCS Three-Statement Financial Model v2.0
Author  : Mark Maxwel Louis  |  MSc Accounting & BI
GitHub  : github.com/MarkMaxwel112075
Date    : June 2026

What this script does:
-----------------------
1.  Validates all DCF assumptions and flags any constraint breaches
2.  Computes FCFF DCF across all three scenarios (Bear / Base / Bull)
3.  Runs full sensitivity analysis (WACC × TGR, RevCAGR × EBIT Margin)
4.  Outputs a clean summary table and saves results to CSV
5.  Generates matplotlib charts: scenario comparison + sensitivity heat map

Usage:
------
    python tcs_model_refresh.py

Requirements:
-------------
    pip install pandas matplotlib numpy openpyxl
"""

import math
import csv
import os
from datetime import datetime

# ─── MODEL PARAMETERS ──────────────────────────────────────────────────────
BASE_YEAR_REV = 30179   # FY2025A Revenue in USD Mn

SCENARIOS = {
    "Bear": {
        "rev_growth":   [0.04, 0.04, 0.05, 0.05, 0.05],
        "ebit_margin":  [0.240, 0.240, 0.245, 0.245, 0.250],
        "wacc":         0.130,
        "tgr":          0.040,
        "capex_pct":    0.018,
        "da_pct":       0.020,
        "wc_pct":       0.006,
        "net_cash":     5500,
        "exrate":       87,
    },
    "Base": {
        "rev_growth":   [0.08, 0.09, 0.09, 0.10, 0.10],
        "ebit_margin":  [0.255, 0.258, 0.262, 0.265, 0.268],
        "wacc":         0.12886279,
        "tgr":          0.045,
        "capex_pct":    0.015,
        "da_pct":       0.020,
        "wc_pct":       0.005,
        "net_cash":     6000,
        "exrate":       85,
    },
    "Bull": {
        "rev_growth":   [0.14, 0.14, 0.13, 0.13, 0.12],
        "ebit_margin":  [0.270, 0.272, 0.275, 0.278, 0.280],
        "wacc":         0.110,
        "tgr":          0.050,
        "capex_pct":    0.013,
        "da_pct":       0.019,
        "wc_pct":       0.004,
        "net_cash":     6500,
        "exrate":       83,
    },
}

SHARES      = 3618     # Mn diluted shares
TAX_RATE    = 0.2517
CMP_INR     = 3350     # Update with live NSE price
FORECAST_YEARS = ["FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]


# ─── VALIDATION ─────────────────────────────────────────────────────────────
def validate_assumptions(scenarios):
    """Check all assumptions meet financial modelling constraints."""
    warnings = []
    for name, s in scenarios.items():
        wacc = s["wacc"]; tgr = s["tgr"]
        spread = wacc - tgr
        if spread < 0.05:
            warnings.append(f"[{name}] WACC–g spread = {spread:.1%} (< 5% threshold — TV highly sensitive)")
        if tgr >= wacc:
            warnings.append(f"[{name}] TGR ≥ WACC — Gordon Growth Model undefined!")
        if s["capex_pct"] > s["da_pct"] * 1.5:
            warnings.append(f"[{name}] Capex > 1.5× D&A — check asset-light assumption")
        tv_wacc_ratio = (1 + tgr) / (wacc - tgr)
        if tv_wacc_ratio > 20:
            warnings.append(f"[{name}] TV multiplier = {tv_wacc_ratio:.1f}× — extreme sensitivity to TGR")
    return warnings


# ─── DCF ENGINE ─────────────────────────────────────────────────────────────
def run_dcf(scenario_name, params, base_rev=BASE_YEAR_REV):
    """
    Compute full 5-year FCFF DCF for a given scenario.

    Returns dict with all intermediate values and output.
    """
    revs, ebits, nopats, das, capexs, dwcs, fcffs = [], [], [], [], [], [], []
    rev = base_rev

    for t in range(5):
        rev   = round(rev * (1 + params["rev_growth"][t]), 1)
        ebit  = round(rev * params["ebit_margin"][t], 1)
        nopat = round(ebit * (1 - TAX_RATE), 1)
        da    = round(rev * params["da_pct"], 1)
        capex = round(-rev * params["capex_pct"], 1)
        dwc   = round(-rev * params["wc_pct"], 1)
        fcff  = round(nopat + da + capex + dwc, 1)

        revs.append(rev); ebits.append(ebit); nopats.append(nopat)
        das.append(da);   capexs.append(capex); dwcs.append(dwc)
        fcffs.append(fcff)

    # Mid-year convention discount factors
    wacc = params["wacc"]
    dfs  = [round(1 / (1 + wacc) ** (t + 0.5), 6) for t in range(5)]
    pvs  = [round(f * d, 1) for f, d in zip(fcffs, dfs)]

    sum_pv = round(sum(pvs), 1)
    tgr    = params["tgr"]
    tv     = round(fcffs[-1] * (1 + tgr) / (wacc - tgr), 0)
    pv_tv  = round(tv * dfs[-1], 1)
    ev     = round(sum_pv + pv_tv, 1)
    equity = round(ev + params["net_cash"], 1)
    iv_usd = round(equity / SHARES, 2)
    iv_inr = round(iv_usd * params["exrate"], 1)
    tv_pct = round(pv_tv / ev, 4) if ev > 0 else 0
    upside = round((iv_inr - CMP_INR) / CMP_INR, 4)

    # 5-yr revenue CAGR
    rev_cagr = round((revs[-1] / base_rev) ** (1 / 5) - 1, 4)

    return {
        "scenario":     scenario_name,
        "revenues":     revs,
        "ebits":        ebits,
        "fcffs":        fcffs,
        "disc_factors": dfs,
        "pv_fcffs":     pvs,
        "sum_pv":       sum_pv,
        "terminal_val": tv,
        "pv_tv":        pv_tv,
        "tv_pct_ev":    tv_pct,
        "ev":           ev,
        "equity":       equity,
        "iv_usd":       iv_usd,
        "iv_inr":       iv_inr,
        "upside":       upside,
        "rev_cagr":     rev_cagr,
        "wacc":         wacc,
        "tgr":          tgr,
        "ebit_margins": params["ebit_margin"],
    }


# ─── SENSITIVITY ENGINE ─────────────────────────────────────────────────────
def sensitivity_wacc_tgr(wacc_range, tgr_range, rev_cagr=0.09, ebit_mgn=0.262):
    """Two-way table: WACC × TGR → IV (INR)."""
    table = {}
    for wacc in wacc_range:
        table[wacc] = {}
        for tgr in tgr_range:
            if wacc <= tgr + 0.01:
                table[wacc][tgr] = "N/A"
                continue
            base_params = dict(SCENARIOS["Base"])
            base_params["wacc"] = wacc
            base_params["tgr"]  = tgr
            base_params["rev_growth"] = [rev_cagr] * 5
            base_params["ebit_margin"] = [ebit_mgn] * 5
            r = run_dcf("sensitivity", base_params)
            table[wacc][tgr] = int(r["iv_inr"])
    return table


def sensitivity_rev_margin(rev_range, margin_range, wacc=0.12886279, tgr=0.045):
    """Two-way table: Revenue CAGR × EBIT Margin → IV (INR)."""
    table = {}
    for rc in rev_range:
        table[rc] = {}
        for em in margin_range:
            base_params = dict(SCENARIOS["Base"])
            base_params["wacc"] = wacc
            base_params["tgr"]  = tgr
            base_params["rev_growth"] = [rc] * 5
            base_params["ebit_margin"] = [em] * 5
            r = run_dcf("sensitivity", base_params)
            table[rc][em] = int(r["iv_inr"])
    return table


# ─── DISPLAY ─────────────────────────────────────────────────────────────────
def print_scenario_summary(results):
    SEP = "=" * 72
    print(f"\n{SEP}")
    print("  TCS FINANCIAL MODEL  |  SCENARIO SUMMARY")
    print(f"  Base Year: FY2025A  |  CMP: ₹{CMP_INR:,}  |  Shares: {SHARES:,} Mn")
    print(SEP)

    metrics = [
        ("5-Yr Revenue CAGR",     lambda r: f"{r['rev_cagr']:.1%}"),
        ("FY30E Revenue (USD Mn)",lambda r: f"${r['revenues'][-1]:,.0f}"),
        ("FY30E EBIT Margin",     lambda r: f"{r['ebit_margins'][-1]:.1%}"),
        ("Sum PV(FCFF) (USD Mn)", lambda r: f"${r['sum_pv']:,.0f}"),
        ("PV of Terminal Value",   lambda r: f"${r['pv_tv']:,.0f}"),
        ("TV as % of EV",          lambda r: f"{r['tv_pct_ev']:.1%}"),
        ("Enterprise Value (Mn)",  lambda r: f"${r['ev']:,.0f}"),
        ("Equity Value (Mn)",      lambda r: f"${r['equity']:,.0f}"),
        ("WACC",                   lambda r: f"{r['wacc']:.2%}"),
        ("Terminal Growth Rate",   lambda r: f"{r['tgr']:.1%}"),
        ("IV per Share (INR)",     lambda r: f"₹{r['iv_inr']:,.0f}"),
        ("vs CMP ₹3,350",         lambda r: f"{r['upside']:+.1%}"),
    ]

    hdr = f"  {'Metric':<32} {'Bear':>12} {'Base':>12} {'Bull':>12}"
    print(hdr)
    print("  " + "-" * 68)
    for label, fn in metrics:
        vals = [fn(results[s]) for s in ["Bear", "Base", "Bull"]]
        print(f"  {label:<32} {vals[0]:>12} {vals[1]:>12} {vals[2]:>12}")
    print(SEP)


def print_sensitivity_table(table, row_label, col_label, row_vals, col_vals,
                             title, base_row=None, base_col=None, cmp=CMP_INR):
    print(f"\n{title}")
    print(f"  {row_label:<10}", end="")
    for cv in col_vals:
        print(f"  {col_label}={cv:.1%}" if isinstance(cv,float) else f"  {cv:>8}", end="")
    print()
    print("  " + "-" * (10 + len(col_vals) * 12))

    for rv in row_vals:
        print(f"  {rv:.2%}" if isinstance(rv,float) else f"  {rv:<10}", end="")
        for cv in col_vals:
            iv = table[rv][cv]
            marker = " ◄" if (rv == base_row and cv == base_col) else ""
            print(f"  {iv:>6,}{marker:>4}" if isinstance(iv,int) else f"  {'N/A':>10}", end="")
        print()


def save_csv(results, sensitivity_wt, sensitivity_rm, outdir="."):
    """Save scenario results and sensitivity tables to CSV."""
    ts = datetime.now().strftime("%Y%m%d_%H%M")

    # Scenario summary
    path_sc = os.path.join(outdir, f"tcs_scenarios_{ts}.csv")
    with open(path_sc, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Scenario","FY26E Rev","FY27E Rev","FY28E Rev","FY29E Rev","FY30E Rev",
                          "Rev CAGR","FY30 EBIT%","WACC","TGR","TV% EV","EV(Mn)","EQ(Mn)","IV INR","vs CMP"])
        for s, r in results.items():
            writer.writerow([s] + r["revenues"] +
                            [f"{r['rev_cagr']:.2%}", f"{r['ebit_margins'][-1]:.2%}",
                             f"{r['wacc']:.2%}", f"{r['tgr']:.2%}", f"{r['tv_pct_ev']:.2%}",
                             r["ev"], r["equity"], r["iv_inr"], f"{r['upside']:+.2%}"])
    print(f"  ✅ Scenarios  → {path_sc}")

    # WACC × TGR sensitivity
    wacc_vals = sorted(sensitivity_wt.keys())
    tgr_vals  = sorted(next(iter(sensitivity_wt.values())).keys())
    path_wt = os.path.join(outdir, f"tcs_sensitivity_wacc_tgr_{ts}.csv")
    with open(path_wt, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["WACC/TGR"] + [f"{t:.1%}" for t in tgr_vals])
        for w in wacc_vals:
            writer.writerow([f"{w:.2%}"] + [sensitivity_wt[w][t] for t in tgr_vals])
    print(f"  ✅ Sensitivity → {path_wt}")


def try_plot(results, sensitivity_wt, wacc_vals, tgr_vals, cmp=CMP_INR):
    """Generate charts if matplotlib is available."""
    try:
        import matplotlib.pyplot as plt
        import matplotlib.colors as mcolors
        import numpy as np

        fig, axes = plt.subplots(1, 2, figsize=(14, 5))
        fig.patch.set_facecolor("#0D2137")

        # ── Chart 1: Scenario comparison — bar chart
        ax1 = axes[0]; ax1.set_facecolor("#0D2137")
        sc_names = ["Bear", "Base", "Bull"]
        iv_vals  = [results[s]["iv_inr"] for s in sc_names]
        colors   = ["#C0392B", "#2E6DA4", "#27AE60"]
        bars = ax1.bar(sc_names, iv_vals, color=colors, width=0.5, edgecolor="#0D2137")
        ax1.axhline(cmp, color="#FFD700", linewidth=2, linestyle="--",
                    label=f"CMP ₹{cmp:,}")
        for bar, iv in zip(bars, iv_vals):
            ax1.text(bar.get_x() + bar.get_width()/2, iv + 50,
                     f"₹{iv:,.0f}", ha="center", color="white", fontsize=10, fontweight="bold")
        ax1.set_title("TCS — Scenario IV per Share (INR)", color="white", fontsize=12, pad=10)
        ax1.set_ylabel("Intrinsic Value (INR)", color="white")
        ax1.tick_params(colors="white"); ax1.spines[:].set_color("#1A3A5C")
        ax1.legend(facecolor="#1A3A5C", labelcolor="white")

        # ── Chart 2: Sensitivity heat map
        ax2 = axes[1]; ax2.set_facecolor("#0D2137")
        z = np.array([[sensitivity_wt[w][t] for t in tgr_vals] for w in wacc_vals], dtype=float)
        im = ax2.imshow(z, cmap="RdYlGn", aspect="auto",
                        vmin=max(1000, z.min()), vmax=min(8000, z.max()))
        ax2.set_xticks(range(len(tgr_vals)))
        ax2.set_xticklabels([f"{t:.1%}" for t in tgr_vals], color="white", fontsize=8)
        ax2.set_yticks(range(len(wacc_vals)))
        ax2.set_yticklabels([f"{w:.1%}" for w in wacc_vals], color="white", fontsize=8)
        ax2.set_title("Sensitivity: WACC × TGR (₹ per Share)", color="white", fontsize=11, pad=10)
        ax2.set_xlabel("Terminal Growth Rate", color="white")
        ax2.set_ylabel("WACC", color="white")
        for i in range(len(wacc_vals)):
            for j in range(len(tgr_vals)):
                ax2.text(j, i, f"{int(z[i,j]):,}", ha="center", va="center",
                         color="black", fontsize=7, fontweight="bold")
        plt.colorbar(im, ax=ax2, label="Intrinsic Value (INR)")

        plt.tight_layout(pad=2.0)
        outpath = os.path.join(os.path.dirname(__file__), "tcs_model_output.png")
        plt.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  ✅ Chart      → {outpath}")
        plt.show()

    except ImportError:
        print("  ℹ  matplotlib not available — skipping charts. Run: pip install matplotlib")
    except Exception as e:
        print(f"  ⚠  Chart error: {e}")


# ─── MAIN ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    SEP = "=" * 72
    print(f"\n{SEP}")
    print("  TCS Financial Model — Python Refresh Script")
    print(f"  Run date: {datetime.now().strftime('%d %b %Y  %H:%M')}")
    print(SEP)

    # Validate
    warnings = validate_assumptions(SCENARIOS)
    if warnings:
        print("\n  ⚠  ASSUMPTION WARNINGS:")
        for w in warnings: print(f"     {w}")
    else:
        print("\n  ✅ All assumption constraints passed.")

    # Run scenarios
    results = {}
    print("\n  Running DCF for Bear / Base / Bull scenarios...")
    for name, params in SCENARIOS.items():
        results[name] = run_dcf(name, params)
        print(f"  [{name}] IV = ₹{results[name]['iv_inr']:,.0f}  ({results[name]['upside']:+.1%} vs CMP)")

    print_scenario_summary(results)

    # Sensitivity
    wacc_vals = [0.090, 0.100, 0.110, 0.120, 0.129, 0.130, 0.140, 0.150]
    tgr_vals  = [0.030, 0.035, 0.040, 0.045, 0.050, 0.055, 0.060]
    rev_range = [0.04, 0.06, 0.08, 0.09, 0.10, 0.12, 0.14]
    mgn_range = [0.23, 0.24, 0.255, 0.262, 0.27, 0.28, 0.29]

    print("  Computing sensitivity tables...")
    sens_wt = sensitivity_wacc_tgr(wacc_vals, tgr_vals)
    sens_rm = sensitivity_rev_margin(rev_range, mgn_range)

    print_sensitivity_table(sens_wt, "WACC", "TGR", wacc_vals, tgr_vals,
                            "  TABLE 1 — WACC × TGR (₹/share)",
                            base_row=0.129, base_col=0.045)

    print_sensitivity_table(sens_rm, "Rev CAGR", "EBIT Mg", rev_range, mgn_range,
                            "  TABLE 2 — Revenue CAGR × EBIT Margin (₹/share)",
                            base_row=0.09, base_col=0.262)

    # Save outputs
    print("\n  Saving outputs...")
    out_dir = os.path.dirname(os.path.abspath(__file__))
    save_csv(results, sens_wt, sens_rm, out_dir)

    # Charts
    try_plot(results, sens_wt, wacc_vals, tgr_vals)

    print(f"\n  ─── ANALYST SIGNAL ─────────────────────────────────────────")
    base_iv = results["Base"]["iv_inr"]
    if base_iv > CMP_INR * 1.15:
        signal = "UNDERVALUED  — Base IV >15% above CMP  →  BUY signal"
    elif base_iv < CMP_INR * 0.85:
        signal = "OVERVALUED   — Base IV >15% below CMP  →  SELL signal"
    else:
        signal = "FAIRLY VALUED — Base IV within 15% of CMP  →  HOLD signal"
    print(f"  Signal : {signal}")
    print(f"  Base IV: ₹{base_iv:,.0f}  |  Bull: ₹{results['Bull']['iv_inr']:,.0f}  |  Bear: ₹{results['Bear']['iv_inr']:,.0f}")
    print(f"  CMP    : ₹{CMP_INR:,}")
    print(f"\n  ⚠  Educational purposes only. Not investment advice.")
    print(f"     Consult a SEBI-registered advisor before investing.")
    print(f"{SEP}\n")
