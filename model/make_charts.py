#!/usr/bin/env python3
"""
Render the three report charts from the model workbook.

  charts/01_history.png          FY20-FY26 sales, NPAT and year-end cash (small multiples, one axis each)
  charts/02_revenue_scenarios.png Revenue paths FY26A-FY32E by operating case
  charts/03_irr_vs_offer.png     Sponsor 5-year IRR vs offer price by operating case

Revenue paths are read by cycling MCASE through LibreOffice (model/run_cases.py helpers);
IRRs are recomputed from the case-capture table on the Sensitivity tab (same cash-flow
definition as the workbook: -[p x FD x (1+fee%) + non-price component], distributions, exit equity).
"""

import json
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import openpyxl  # noqa: E402

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
XLSX = HERE / "Project_Daylight_LBO_v1.xlsx"
OUT = ROOT / "charts"
OUT.mkdir(exist_ok=True)

# Reference palette (light mode) - categorical slots in fixed order; validated with
# validate_palette.js (adjacent CVD dE >= 9.1). Slots 3-5 are < 3:1 on the surface, so every
# series is direct-labelled.
SURFACE, INK, INK2, MUTED, GRID, AXIS = "#fcfcfb", "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7"
SLOTS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
CASES = ["Downside", "Conservative", "Sponsor base", "Orals underwhelm", "Upside (vitiligo)"]
PROBS = [0.20, 0.30, 0.30, 0.15, 0.05]

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": AXIS, "axes.labelcolor": INK2,
    "xtick.color": MUTED, "ytick.color": MUTED, "axes.facecolor": SURFACE, "figure.facecolor": SURFACE,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "grid.linestyle": "-",
    "axes.spines.top": False, "axes.spines.right": False, "axes.titleweight": "bold", "axes.titlesize": 10,
    "axes.titlecolor": INK, "legend.frameon": False,
})


def style_axes(ax):
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="both", length=0)
    ax.set_axisbelow(True)


def irr(flows, lo=-0.99, hi=3.0):
    def npv(r):
        return sum(cf / (1 + r) ** t for t, cf in enumerate(flows))
    f_lo, f_hi = npv(lo), npv(hi)
    if f_lo * f_hi > 0:
        return float("nan")
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if f_lo * f_mid <= 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


def revenue_paths():
    """Cycle MCASE 1-5 in LibreOffice and read the revenue row (cached in charts/_case_paths.json)."""
    cache = OUT / "_case_paths.json"
    if cache.exists() and cache.stat().st_mtime > XLSX.stat().st_mtime:
        return json.loads(cache.read_text())
    sys.path.insert(0, str(HERE))
    import shutil
    import tempfile
    import time

    import run_cases as rc
    wbf = openpyxl.load_workbook(XLSX)
    lbo = wbf["LBO"]
    rev_row = next(r for r in range(1, lbo.max_row + 1) if lbo[f"B{r}"].value == "Total revenue")
    tmp = Path(tempfile.mkdtemp()) / "paths.xlsx"
    shutil.copy(XLSX, tmp)
    proc, ctx = rc.start_office(port=2099)
    paths = {}
    desktop = None
    try:
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(tmp.as_uri(), "_blank", 0,
                                           (rc._prop("Hidden", True), rc._prop("FilterName", "Calc MS Excel 2007 XML")))
        doc.IsIterationEnabled = True
        doc.IterationCount = 1000
        doc.IterationEpsilon = 1e-9
        mc = doc.NamedRanges.getByName("MCASE").getReferredCells().getCellByPosition(0, 0)
        sheet = doc.Sheets.getByName("LBO")
        for n in range(1, 6):
            mc.setValue(n)
            rc.converge(doc)
            paths[str(n)] = [sheet.getCellRangeByName(f"{c}{rev_row}").getValue() for c in "GHIJKLM"]
        doc.close(True)
    finally:
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
    cache.write_text(json.dumps(paths))
    return paths


def captured_cases():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["Sensitivity"]
    hdr = next(r for r in range(1, ws.max_row + 1) if ws[f"B{r}"].value == "Case" and ws[f"C{r}"].value == "Label")
    cols = {ws.cell(row=hdr, column=c).value: c for c in range(4, 40) if ws.cell(row=hdr, column=c).value}
    out = {}
    for i in range(1, 6):
        r = hdr + i
        get = lambda lab: ws.cell(row=r, column=cols[lab]).value
        out[i] = {"npi": get("Non-price inv. component"), "eqx": get("Exit equity (Y5)"),
                  "d": [get(f"Dist Y{t}") for t in range(1, 6)]}
    lbo = wb["LBO"]
    fd = next(lbo[f"D{r}"].value for r in range(1, 80) if lbo[f"B{r}"].value == "Fully diluted shares (m)")
    fee = next(lbo[f"D{r}"].value for r in range(1, 80) if lbo[f"B{r}"].value == "Transaction expenses (% of equity purchase price)")
    und = next(lbo[f"D{r}"].value for r in range(1, 80) if lbo[f"B{r}"].value == "Undisturbed share price (A$)")
    return out, fd, fee, und


def chart_history():
    years = ["FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26"]
    sales = [32.6, 48.5, 67.0, 77.9, 88.0, 95.0, 94.0]   # FY20-22 as-reported revenue; FY23 derived
    npat = [15.1, 24.7, 20.9, 30.6, 35.6, 36.2, 33.9]
    cash_years = ["FY21", "FY22", "FY23", "FY24", "FY25", "FY26"]
    cash = [82.7, 121.5, 157.0, 183.7, 224.1, 252.1]
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(10, 4.2), gridspec_kw={"width_ratios": [1.25, 1]})
    x = range(len(years))
    a1.plot(x, sales, color=SLOTS[0], lw=2, marker="o", ms=5, mec=SURFACE, mew=1.5, label="Sales revenue")
    a1.plot(x, npat, color=SLOTS[1], lw=2, marker="o", ms=5, mec=SURFACE, mew=1.5, label="Net profit after tax")
    a1.set_xticks(list(x), years)
    a1.set_ylim(0, 110)
    a1.set_title("Sales stalled in FY26 while profit held up (A$m)", loc="left")
    a1.annotate(f"A${sales[-1]:.0f}m", (6, sales[-1]), xytext=(6, 8), textcoords="offset points", ha="center", color=INK2, fontsize=8)
    a1.annotate(f"A${npat[-1]:.0f}m", (6, npat[-1]), xytext=(6, -14), textcoords="offset points", ha="center", color=INK2, fontsize=8)
    a1.legend(loc="upper left", fontsize=8)
    style_axes(a1)
    xb = range(len(cash_years))
    a2.bar(xb, cash, color=SLOTS[0], width=0.62)
    a2.set_xticks(list(xb), cash_years)
    a2.set_ylim(0, 290)
    a2.set_title("Cash pile, no debt (A$m, 30 June)", loc="left")
    for i in (0, len(cash) - 1):
        a2.annotate(f"A${cash[i]:.0f}m", (i, cash[i]), xytext=(0, 4), textcoords="offset points", ha="center", color=INK2, fontsize=8)
    style_axes(a2)
    fig.text(0.01, 0.01, "Source: Clinuvel results releases FY21-FY26; FY20-FY22 sales = reported revenue (interest immaterial); "
             "FY23 sales and FY21 cash derived from reported growth rates.", fontsize=7, color=MUTED)
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    fig.savefig(OUT / "01_history.png", dpi=200)
    plt.close(fig)


def chart_scenarios(paths):
    years = ["FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E", "FY32E"]
    fig, ax = plt.subplots(figsize=(9, 4.8))
    x = list(range(len(years)))
    ends = []
    for i, name in enumerate(CASES):
        ys = paths[str(i + 1)]
        ax.plot(x, ys, color=SLOTS[i], lw=2, label=name)
        ax.plot(x[-1], ys[-1], marker="o", ms=6, color=SLOTS[i], mec=SURFACE, mew=1.5)
        ends.append((ys[-1], name, i))
    # direct labels at line ends, nudged apart
    ends.sort()
    placed = []
    for y, name, i in ends:
        yy = y
        for p in placed:
            if abs(yy - p) < 6:
                yy = p + 6
        placed.append(yy)
        ax.annotate(f"{name}  A${y:.0f}m", (x[-1], y), xytext=(10, yy - y), textcoords="offset points",
                    va="center", fontsize=8, color=INK2)
    ax.axvspan(0.5, 1.5, color=GRID, alpha=0.45, lw=0)
    ax.text(1.0, 8, "Oral competitor\nFDA decisions\n(Feb & mid-2027)", ha="center", va="bottom", fontsize=7, color=INK2)
    ax.set_xticks(x, years)
    ax.set_xlim(-0.2, len(years) + 1.6)
    ax.set_ylim(0, 155)
    ax.set_title("Revenue by operating case (A$m): the base case already assumes an oral competitor", loc="left")
    ax.legend(loc="lower center", fontsize=8, ncol=2, bbox_to_anchor=(0.56, 0.02))
    style_axes(ax)
    fig.text(0.01, 0.01, "Source: Project Daylight model (OPCASE). Upside includes vitiligo revenue from FY30. "
             "US share of FY26 sales assumed 40% (not disclosed in accessible sources).", fontsize=7, color=MUTED)
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    fig.savefig(OUT / "02_revenue_scenarios.png", dpi=200)
    plt.close(fig)


def chart_irr(cases, fd, fee, und):
    prices = [8.25 + 0.25 * i for i in range(20)]  # 8.25 .. 13.00
    fig, ax = plt.subplots(figsize=(9, 5))
    series = {}
    for i in range(1, 6):
        c = cases[i]
        ys = []
        for p in prices:
            inv = p * fd * (1 + fee) + c["npi"]
            flows = [-inv] + c["d"][:4] + [c["d"][4] + c["eqx"]]
            ys.append(irr(flows) * 100)
        series[i] = ys
        ax.plot(prices, ys, color=SLOTS[i - 1], lw=2, label=CASES[i - 1])
    exp = []
    for p in prices:
        inv = sum(PROBS[i - 1] * (p * fd * (1 + fee) + cases[i]["npi"]) for i in range(1, 6))
        flows = [-inv] + [sum(PROBS[i - 1] * cases[i]["d"][t] for i in range(1, 6)) for t in range(4)]
        flows.append(sum(PROBS[i - 1] * (cases[i]["d"][4] + cases[i]["eqx"]) for i in range(1, 6)))
        exp.append(irr(flows) * 100)
    ax.plot(prices, exp, color=INK, lw=2.6, label="Probability-weighted cash flows")
    ax.axhline(20, color=INK2, lw=1, ls=(0, (4, 3)))
    ax.text(9.45, 21, "20% sponsor hurdle", ha="left", va="bottom", fontsize=8, color=INK2)
    ax.axvline(und, color=MUTED, lw=1)
    ax.text(und + 0.05, -26, f"Undisturbed\nA${und:.2f}", fontsize=7.5, color=INK2, va="bottom")
    ax.axvline(10.75, color=MUTED, lw=1)
    ax.text(10.8, -26, "Base offer\nA$10.75 (+30%)", fontsize=7.5, color=INK2, va="bottom")
    # direct labels at the right edge
    ends = sorted([(series[i][-1], CASES[i - 1]) for i in range(1, 6)] + [(exp[-1], "Probability-weighted")])
    placed = []
    for y, name in ends:
        yy = y
        for p in placed:
            if abs(yy - p) < 3.2:
                yy = p + 3.2
        placed.append(yy)
        ax.annotate(f"{name} {0 if abs(y) < 0.5 else y:.0f}%", (prices[-1], y), xytext=(8, (yy - y) * 3.2), textcoords="offset points",
                    va="center", fontsize=8, color=INK2)
    # where the expected line crosses the hurdle
    cross = None
    for (p0, y0), (p1, y1) in zip(zip(prices, exp), zip(prices[1:], exp[1:])):
        if y0 >= 20 > y1:
            cross = p0 + (p1 - p0) * (y0 - 20) / (y0 - y1)
            break
    ax.set_xlim(prices[0] - 0.1, prices[-1] + 2.2)
    ax.set_xticks([8.5, 9, 9.5, 10, 10.5, 11, 11.5, 12, 12.5, 13])
    ax.set_ylim(-28, 60)
    ax.set_xlabel("Offer price per share (A$)")
    ax.set_ylabel("Sponsor 5-year IRR (%)")
    ax.set_title("A 20% IRR needs a price close to where Clinuvel already trades", loc="left")
    ax.legend(loc="upper right", fontsize=8, bbox_to_anchor=(0.80, 1.0))
    style_axes(ax)
    fig.text(0.01, 0.012, "Source: Project Daylight model, base financing (2.0x unitranche), close 30-Jun-2027, exit FY32 at case "
             "multiples (3.5x-9.0x); probability weights 20/30/30/15/5%.\n" + (f"The probability-weighted IRR falls below 20% at offers above ~A${cross:.2f} per share." if cross else ""),
             fontsize=7, color=MUTED)
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    fig.savefig(OUT / "03_irr_vs_offer.png", dpi=200)
    plt.close(fig)
    return prices, series, exp


if __name__ == "__main__":
    chart_history()
    chart_scenarios(revenue_paths())
    cases, fd, fee, und = captured_cases()
    prices, series, exp = chart_irr(cases, fd, fee, und)
    for p, e in zip(prices, exp):
        print(f"A${p:5.2f}  expected IRR {e:5.1f}%  " + "  ".join(f"{CASES[i-1][:10]} {series[i][prices.index(p)]:5.1f}%" for i in range(1, 6)))
    print("charts written to", OUT)
