#!/usr/bin/env python3
"""
Build "Project Daylight": a take-private screening LBO for Clinuvel Pharmaceuticals
(ASX: CUV / Nasdaq: CUVL).

Architecture (Multiple Expansion conventions):
    Cover | Hist | MCASE | OPCASE | FINCASE | LBO | ATP | Contrib | Sensitivity | CVR

* MCASE drives everything: operating case, financing case, offer price, fees, rollover,
  exit multiple and minimum cash.  OPCASE / FINCASE use the OFFSET x step pattern.
* CIRC switch (named) toggles average-balance interest (iterative calculation on).
* Sensitivity case tables use self-referencing IF capture cells; populate them by running
  every MCASE (model/run_cases.py does this headlessly in LibreOffice).

Colour code: blue = hard-coded input, black = same-sheet formula, green = cross-sheet link.
All A$ in millions unless stated.  FYE 30 June.

Run:  python3 model/build_model.py && python3 model/run_cases.py model/Project_Daylight_LBO_v1.xlsx
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties

OUT = Path(__file__).resolve().parent / "Project_Daylight_LBO_v1.xlsx"

# ----------------------------------------------------------------------------------------
# Style constants
# ----------------------------------------------------------------------------------------
FONT = "Arial"
BLUE, BLACK, GREEN, RED, WHITE, GRAY, NAVY = "0000FF", "000000", "008000", "C00000", "FFFFFF", "7F7F7F", "1F3864"
F_TITLE = PatternFill("solid", start_color=NAVY)
F_HDR = PatternFill("solid", start_color="DDEBF7")
F_OUT = PatternFill("solid", start_color="595959")
F_RUN = PatternFill("solid", start_color="EDEDED")
F_KEY = PatternFill("solid", start_color="FFF2CC")
F_CHECK = PatternFill("solid", start_color="E2EFDA")
THIN = Side(style="thin", color="A6A6A6")
MED = Side(style="medium", color="000000")
BOX = Border(left=MED, right=MED, top=MED, bottom=MED)
TOPLINE = Border(top=Side(style="thin", color="000000"))

M1 = '#,##0.0;(#,##0.0);"-"'
M0 = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
MULT = '0.0"x";(0.0"x");"-"'
MULT2 = '0.00"x";(0.00"x");"-"'
PS = '"A$"0.00;("A$"0.00);"-"'
SH = '#,##0.00'
DATEF = 'dd-mmm-yy'
INTF = '0'

YEARS = ["FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E", "FY32E"]
YC = ["G", "H", "I", "J", "K", "L", "M"]          # LBO year columns
PROJ = ["I", "J", "K", "L", "M"]                 # sponsor-ownership years 1..5
H_TO_M = ["H", "I", "J", "K", "L", "M"]          # FY27E..FY32E (OPCASE columns align)
PREV = {c: YC[i - 1] for i, c in enumerate(YC) if i > 0}


def col_shift(c, n):
    return get_column_letter(column_index_from_string(c) + n)


def auto_color(v):
    if isinstance(v, str) and v.startswith("="):
        return GREEN if "!" in v else BLACK
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return BLUE
    return BLACK


def put(ws, ref, v, fmt=None, bold=False, color=None, fill=None, align=None, italic=False,
        border=None, wrap=False, size=9):
    c = ws[ref]
    c.value = v
    c.font = Font(name=FONT, size=size, color=color or auto_color(v), bold=bold, italic=italic)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="top" if wrap else "center", wrap_text=wrap)
    if border:
        c.border = border
    return c


def bar(ws, row, text, c1="B", c2="O", fill=F_HDR, color=NAVY, size=9):
    for ci in range(column_index_from_string(c1), column_index_from_string(c2) + 1):
        ws.cell(row=row, column=ci).fill = fill
    put(ws, f"{c1}{row}", text, bold=True, color=color, fill=fill, size=size)


def title(ws, text, sub, c2="O"):
    bar(ws, 2, text, "B", c2, fill=F_TITLE, color=WHITE, size=12)
    put(ws, "B3", sub, italic=True, color=GRAY, size=8)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def irr_f(rng):
    """IRR with fallback guesses (deeply negative outcomes do not converge from a 15% guess)."""
    return f"IFERROR(IRR({rng},0.15),IFERROR(IRR({rng},-0.2),IFERROR(IRR({rng},-0.6),-1)))"


def set_rows(ws, height=12.75):
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = height


# ----------------------------------------------------------------------------------------
# INPUTS (every hard-code lives here, with its source)
# ----------------------------------------------------------------------------------------
SRC_FY26 = "FY26 results release, 27-Aug-2026 (GlobeNewswire/Nasdaq; 6-K)"
FY26 = {
    "product_rev": (94.024, "Revenue from ordinary activities; " + SRC_FY26),
    "total_rev": (101.153, "Total revenues incl. interest & other income; " + SRC_FY26),
    "expenses": (53.5, "Total expenses; " + SRC_FY26),
    "npat": (33.9, "NPAT (-6% y/y); " + SRC_FY26),
    "cash": (252.055, "Cash reserves (+12%); " + SRC_FY26),
    "net_assets": (272.9, "Net assets 30-Jun-2026; strategic reorganisation release 23-Jul-2026"),
    "ocf": (36.9, "Operating cash inflow; " + SRC_FY26),
    "us_share": (0.40, "ASSUMPTION - US share of FY26 product revenue not in accessible disclosure; test 30-50%"),
    "nwc": (15.0, "ESTIMATE - net working capital & other net operating items (c.16% of sales)"),
}
# FY26 expense split: ESTIMATES that sum to the reported A$53.5m (not disclosed by function).
COSTS = [
    ("cogs", "Cost of goods / supply & distribution", 7.0, "EST: gross margin >90% (Kalkine, 2026)"),
    ("core", "Core SCENESSE commercial, medical, PV & regulatory", 11.0, "EST: A$8-11m range (research stream)"),
    ("corp", "Corporate, listing & G&A", 10.5, "EST: A$8-11m range incl. ASX/SEC listing costs"),
    ("cbm", "Communications, branding, marketing & consumer (CBM)", 5.0, "EST: CBM +100% in FY25; A$4-7m"),
    ("vit", "Vitiligo clinical program (CUV105/CUV107)", 12.0, "EST: A$10-15m; clinical dev. +215% in FY25"),
    ("rd", "Other R&D (Singapore, PRENUMBRA, NEURACTHEL, XP/VP)", 6.0, "EST: A$6-10m"),
    ("da", "Depreciation & amortisation", 2.0, "EST"),
]
HIST = {  # FY20..FY26; None = not verified in accessible sources
    "years": ["FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26"],
    "product_rev": [None, None, None, 77.9, 88.0, 95.018, 94.024],
    "total_rev": [32.565, 48.451, 66.987, 83.0, 95.31, 105.3, 101.153],
    "expenses": [None, None, None, None, 44.75, 53.7, 53.5],
    "pbt": [None, None, 34.321, None, 50.6, 51.6, 47.653],
    "npat": [15.1, 24.728, 20.876, 30.605, 35.64, 36.2, 33.9],
    "cash": [None, None, 121.509, 157.0, 183.7, 224.1, 252.055],
    "dps": [None, None, 0.04, None, None, 0.05, 0.05],
}
HIST_NOTES = {
    "product_rev": "FY23 derived (FY24 +13% to A$88m, Morningstar); FY25 A$95.018m (Appendix 4E); FY26 A$94.024m",
    "total_rev": "As reported; FY20/FY21 'revenue' as reported (interest immaterial then)",
    "expenses": "FY24 derived from FY25 +20%; FY25/FY26 as reported",
    "pbt": "FY22 as reported (adj. PBT A$39.8m); FY24 derived (FY25 +2%); FY26 derived",
    "npat": "FY20 derived from FY21 +64%; FY22 as reported (non-cash items); others as reported",
    "cash": "FY24 derived (FY25 +22%); others as reported",
    "dps": "Fully franked where stated",
}

TXN = {
    "undisturbed": (8.24, "ASX close 1-Sep-2026 (A$9.18 on 27-Aug results day, -12.7%); CUVL US$5.85 on 16-Sep"),
    "shares_basic": (50.43, "Derived: Vanguard 2,523,018 sh = 5.003% (Form 603, 14-Aug-2026)"),
    "dilutive": (0.25, "EST: performance rights; CEO holds no LTI (Jul-2026 terms)"),
    "debt_existing": (0.0, "No borrowings (leases ignored)"),
    "exit_cost_pct": (0.015, "Sell-side fees at exit, % of TEV"),
    "tax_rate": (0.28, "FY25 29.8%, FY26 28.9% effective; US redomicile may lower"),
    "deposit_rate": (0.0325, "Yield on cash (H1 FY26 interest A$5.3m on ~A$230m)"),
    "dps_fy27": (0.05, "FY26 final dividend paid 18-Sep-2026 (fully franked)"),
    "da_run": (2.0, "D&A run-rate (EST)"),
    "us_incr": (2.5, "EST: incremental US HQ / Nasdaq / SEC run-rate cost from FY27"),
    "dist_pct": (1.0, "Share of excess cash distributed to equity once debt repaid"),
    "min_debt_incr": (5.0, "Debt sizing rounding increment"),
}
BASE_RATE = [0.036, 0.036, 0.036, 0.036, 0.036]  # FY28..FY32, BBSY/SOFR proxy (assumption)

# ---------------------------------- operating cases -------------------------------------
# Columns: FY27E (status quo, pre-close) then FY28E..FY32E (sponsor plan).
OP_DRIVERS = [
    ("us_g", "US SCENESSE revenue growth", "%", PCT),
    ("eu_g", "Europe & RoW SCENESSE revenue growth", "%", PCT),
    ("vit_rev", "Vitiligo revenue", "A$m", M1),
    ("oth_rev", "Other new product revenue", "A$m", M1),
    ("cogs", "COGS % of revenue", "%", PCT),
    ("infl", "Core & corporate cost inflation", "%", PCT),
    ("def", "Competitive-defence spend (patient services, access)", "A$m", M1),
    ("sav", "Sponsor corporate cost savings (% of status quo)", "%", PCT),
    ("cbm", "CBM & consumer spend", "A$m", M1),
    ("vit", "Vitiligo clinical spend", "A$m", M1),
    ("rd", "Other R&D spend", "A$m", M1),
    ("vitl", "Vitiligo launch & commercial costs", "A$m", M1),
    ("nonrec", "Non-recurring items (relocation, restructuring)", "A$m", M1),
    ("capex", "Capex", "A$m", M1),
    ("nwc", "Net working capital % of revenue", "%", PCT),
]
SQ = {  # FY27E status-quo (company plan) values common to all cases
    "cogs": 0.075, "infl": 0.03, "def": 0.0, "sav": 0.0, "cbm": -5.2, "vit": -15.0, "rd": -6.5,
    "vitl": 0.0, "nonrec": -4.0, "capex": -4.0, "nwc": 0.16, "vit_rev": 0.0, "oth_rev": 0.0,
}


def case(us, eu, fy28_on):
    d = {"us_g": us, "eu_g": eu}
    for k, v in SQ.items():
        d[k] = [v] + fy28_on.get(k, [v] * 5)
    return d


LEAN = {"infl": [0.03] * 5, "sav": [0.35, 0.45, 0.45, 0.45, 0.45], "cbm": [-1.0] * 5,
        "rd": [-1.5] * 5, "vit": [-2.0, 0, 0, 0, 0], "vitl": [0] * 5,
        "nonrec": [-4.0, 0, 0, 0, 0], "capex": [-1.5] * 5, "nwc": [0.16] * 5,
        "vit_rev": [0] * 5, "oth_rev": [0] * 5}
OPCASES = [
    (1, "Downside",
     "Both orals approved (dersimelagon ~Feb-27, bitopertin ~mid-27) and adopted fast; EU oral from FY29; price pressure",
     case([-0.08, -0.35, -0.35, -0.25, -0.20, -0.15], [0.04, 0.02, -0.05, -0.15, -0.20, -0.15],
          {**LEAN, "cogs": [0.08, 0.085, 0.09, 0.095, 0.10], "def": [-2.0, -2.0, -1.5, -1.0, -1.0]})),
    (2, "Conservative",
     "Dersimelagon approved Feb-27 with solid uptake; bitopertin also approved; EU competition from FY30",
     case([-0.06, -0.25, -0.25, -0.15, -0.10, -0.08], [0.04, 0.03, 0.00, -0.08, -0.10, -0.08],
          {**LEAN, "cogs": [0.0775, 0.08, 0.085, 0.085, 0.085], "def": [-2.0, -2.0, -1.5, -1.0, -1.0]})),
    (3, "Sponsor base",
     "Dersimelagon approved; APOLLO misses/delays bitopertin; SCENESSE keeps implant-preferring & severe patients",
     case([-0.05, -0.18, -0.15, -0.08, -0.05, -0.03], [0.05, 0.04, 0.02, -0.03, -0.05, -0.03],
          {**LEAN, "cogs": [0.075] * 5, "def": [-1.5, -1.5, -1.0, -1.0, -1.0]})),
    (4, "Orals underwhelm",
     "Orals approved but modest efficacy; combination use; Canada, adolescents & EU year-round dosing drive growth",
     case([-0.03, -0.05, -0.03, 0.00, 0.02, 0.02], [0.05, 0.05, 0.04, 0.03, 0.02, 0.02],
          {**LEAN, "cogs": [0.075] * 5, "def": [-1.0, -1.0, -0.5, -0.5, -0.5], "rd": [-2.5] * 5})),
    (5, "Upside (vitiligo)",
     "As case 4 plus CUV105 positive: sponsor funds CUV107; vitiligo approval & launch from FY30",
     case([-0.03, -0.05, -0.03, 0.00, 0.02, 0.02], [0.05, 0.05, 0.04, 0.03, 0.02, 0.02],
          {**LEAN, "cogs": [0.075] * 5, "def": [-1.0, -1.0, -0.5, -0.5, -0.5], "rd": [-2.5] * 5,
           "vit": [-14.0, -12.0, -6.0, -3.0, -3.0], "vitl": [0.0, -5.0, -12.0, -15.0, -16.0],
           "vit_rev": [0.0, 0.0, 5.0, 20.0, 40.0]})),
    (6, "Base revenue, no cost-out",
     "Case 3 revenue, but management's plan continues: CUV107 funded, US build-out, CBM & Singapore R&D; no sponsor savings",
     case([-0.05, -0.18, -0.15, -0.08, -0.05, -0.03], [0.05, 0.04, 0.02, -0.03, -0.05, -0.03],
          {"cogs": [0.075] * 5, "infl": [0.03] * 5, "def": [-1.5, -1.5, -1.0, -1.0, -1.0], "sav": [0.0] * 5,
           "cbm": [-5.4, -5.5, -5.7, -5.9, -6.0], "vit": [-15.0, -14.0, -8.0, -5.0, -4.0],
           "rd": [-6.7, -6.9, -7.1, -7.3, -7.5], "vitl": [0.0] * 5, "nonrec": [0.0] * 5, "capex": [-3.0] * 5,
           "nwc": [0.16] * 5, "vit_rev": [0.0] * 5, "oth_rev": [0.0] * 5})),
]
CASE_PROB = [0.20, 0.30, 0.30, 0.15, 0.05]  # illustrative probability weights for cases 1-5
EXIT_MULT = {1: 3.5, 2: 4.5, 3: 5.5, 4: 7.0, 5: 9.0, 6: 5.5}
CORE_OPS = OPCASES[:5]  # the five cases used in the op x fin grids

# ---------------------------------- financing cases -------------------------------------
FIN_DRIVERS = [
    ("tl_lev", "Unitranche term loan (x FY27E Adj. EBITDA)", MULT),
    ("tl_margin", "Term loan margin over base rate", PCT2),
    ("tl_floor", "Term loan base-rate floor", PCT2),
    ("tl_fee", "Term loan upfront fee / OID", PCT2),
    ("tl_amort", "Term loan mandatory amortisation (% p.a.)", PCT),
    ("tl_tenor", "Term loan tenor (years)", INTF),
    ("pik_lev", "HoldCo PIK notes (x FY27E Adj. EBITDA)", MULT),
    ("pik_rate", "PIK coupon", PCT2),
    ("pik_fee", "PIK upfront fee", PCT2),
    ("pik_tenor", "PIK tenor (years)", INTF),
    ("rcf_commit", "Revolver commitment (A$m)", M1),
    ("rcf_margin", "Revolver margin", PCT2),
    ("rcf_cfee", "Revolver commitment fee (on undrawn)", PCT2),
    ("rcf_fee", "Revolver upfront fee", PCT2),
    ("rcf_tenor", "Revolver tenor (years)", INTF),
    ("cov_lev", "Covenant: max net debt / EBITDA", MULT),
    ("cov_icr", "Covenant: min EBITDA / cash interest", MULT),
]
FINCASES = [
    (1, "Base (specialist private credit)", [2.0, 0.060, 0.0, 0.025, 0.05, 5, 0.0, 0.12, 0.02, 6, 20.0, 0.040, 0.014, 0.01, 5, 3.0, 3.0]),
    (2, "High leverage (+ HoldCo PIK)", [3.0, 0.0675, 0.0, 0.030, 0.05, 5, 0.75, 0.12, 0.02, 6, 20.0, 0.040, 0.014, 0.01, 5, 4.0, 2.0]),
    (3, "Low leverage", [1.0, 0.055, 0.0, 0.0225, 0.05, 5, 0.0, 0.12, 0.02, 6, 20.0, 0.040, 0.014, 0.01, 5, 2.5, 4.0]),
    (4, "No acquisition debt (cash-box)", [0.0, 0.060, 0.0, 0.025, 0.05, 5, 0.0, 0.12, 0.02, 6, 20.0, 0.040, 0.014, 0.01, 5, 3.0, 3.0]),
]

# ---------------------------------- master cases ----------------------------------------
BASE_OFFER = 10.75
TRANS_PCT = 0.0275
MIN_CASH = 40.0
MCASES = []
for fin in (1, 2, 3, 4):
    for op in (1, 2, 3, 4, 5):
        n = (fin - 1) * 5 + op
        MCASES.append((n, f"{OPCASES[op-1][1]} | {FINCASES[fin-1][1].split(' (')[0]}", op, fin, BASE_OFFER,
                       TRANS_PCT, 0.0, EXIT_MULT[op], MIN_CASH))
MCASES += [
    (21, "Sponsor base | Base | CEO rolls 6.8%", 3, 1, BASE_OFFER, TRANS_PCT, 0.068, EXIT_MULT[3], MIN_CASH),
    (22, "Sponsor base | Base | A$9.50 offer", 3, 1, 9.50, TRANS_PCT, 0.0, EXIT_MULT[3], MIN_CASH),
    (23, "Sponsor base | Base | A$12.50 offer", 3, 1, 12.50, TRANS_PCT, 0.0, EXIT_MULT[3], MIN_CASH),
    (24, "Base revenue, no cost-out | Base", 6, 1, BASE_OFFER, TRANS_PCT, 0.0, EXIT_MULT[6], MIN_CASH),
]
DEFAULT_CASE = 3

# ========================================================================================
wb = Workbook()
wb.remove(wb.active)
S = {name: wb.create_sheet(name) for name in
     ["Cover", "Hist", "MCASE", "OPCASE", "FINCASE", "LBO", "ATP", "Contrib", "Sensitivity", "CVR"]}
for n, color in [("Cover", "808080"), ("Hist", "808080"), ("MCASE", "A9D08E"), ("OPCASE", "A9D08E"),
                 ("FINCASE", "A9D08E"), ("LBO", NAVY), ("ATP", "F4B084"), ("Contrib", "F4B084"),
                 ("Sensitivity", "F4B084"), ("CVR", "F4B084")]:
    S[n].sheet_properties.tabColor = color
    S[n].sheet_view.showGridLines = False


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ========================================================================================
# Hist
# ========================================================================================
ws = S["Hist"]
title(ws, "Historical financials & FY26A base-year inputs",
      "A$m. Blue = reported or estimated input (see source column). Estimates flagged 'EST'/'ASSUMPTION'.", "L")
widths(ws, {"A": 2, "B": 50, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 2, "K": 70})
bar(ws, 5, "Reported history (FY ending 30 June)", "B", "K")
for i, y in enumerate(HIST["years"]):
    put(ws, f"{get_column_letter(3 + i)}6", y, bold=True, align="right")
put(ws, "K6", "Source / note", bold=True)
hist_rows = [("product_rev", "Product revenue (SCENESSE)", M1), ("total_rev", "Total revenues incl. interest & other income", M1),
             ("expenses", "Total expenses", M1), ("pbt", "Profit before tax", M1), ("npat", "Net profit after tax", M1),
             ("cash", "Cash at year end", M1), ("dps", "Dividend per share (A$)", '0.000')]
HR = {}
r = 7
for key, label, fmt in hist_rows:
    HR[key] = r
    put(ws, f"B{r}", label)
    for i, v in enumerate(HIST[key]):
        if v is not None:
            put(ws, f"{get_column_letter(3 + i)}{r}", v, fmt=fmt)
        else:
            put(ws, f"{get_column_letter(3 + i)}{r}", "n.v.", color=GRAY, align="right")
    put(ws, f"K{r}", HIST_NOTES[key], italic=True, color=GRAY, size=8)
    r += 1
put(ws, f"B{r}", "Product revenue growth", italic=True)
for i in range(1, 7):
    c, p = get_column_letter(3 + i), get_column_letter(2 + i)
    put(ws, f"{c}{r}", f'=IFERROR({c}{HR["product_rev"]}/{p}{HR["product_rev"]}-1,"")', fmt=PCT, italic=True)
r += 1
put(ws, f"B{r}", "NPAT margin on total revenues", italic=True)
for i in range(7):
    c = get_column_letter(3 + i)
    put(ws, f"{c}{r}", f'=IFERROR({c}{HR["npat"]}/{c}{HR["total_rev"]},"")', fmt=PCT, italic=True)
r += 1
put(ws, f"B{r}", "Half-year data: H1 FY26 sales A$36.93m (+4%), NPAT A$10.44m (-26%), cash A$233.0m; H1 interest A$5.26m "
                 "(Appendix 4D, 26-Feb-2026). Implied H2 FY26 sales A$57.1m (c.-4% y/y).", italic=True, color=GRAY, size=8)

r += 2
bar(ws, r, "FY26A base-year inputs used by the LBO tab", "B", "K")
r += 1
put(ws, f"B{r}", "Item", bold=True); put(ws, f"C{r}", "A$m", bold=True, align="right"); put(ws, f"K{r}", "Source / note", bold=True)
r += 1
HI = {}


def hist_input(key, label, val, note, fmt=M1):
    global r
    HI[key] = r
    put(ws, f"B{r}", label)
    put(ws, f"C{r}", val, fmt=fmt)
    put(ws, f"K{r}", note, italic=True, color=GRAY, size=8)
    r += 1


hist_input("product_rev", "Product revenue FY26A", FY26["product_rev"][0], FY26["product_rev"][1])
hist_input("us_share", "US share of product revenue", FY26["us_share"][0], FY26["us_share"][1], PCT)
HI["us_rev"] = r
put(ws, f"B{r}", "  US SCENESSE revenue FY26A"); put(ws, f"C{r}", f"=C{HI['product_rev']}*C{HI['us_share']}", fmt=M1); r += 1
HI["eu_rev"] = r
put(ws, f"B{r}", "  Europe & RoW SCENESSE revenue FY26A"); put(ws, f"C{r}", f"=C{HI['product_rev']}-C{HI['us_rev']}", fmt=M1); r += 1
hist_input("total_rev", "Total revenues incl. interest & other income", FY26["total_rev"][0], FY26["total_rev"][1])
HI["other_inc"] = r
put(ws, f"B{r}", "  Interest & other income (derived)"); put(ws, f"C{r}", f"=C{HI['total_rev']}-C{HI['product_rev']}", fmt=M1); r += 1
hist_input("expenses", "Total expenses (reported)", FY26["expenses"][0], FY26["expenses"][1])
for key, label, val, note in COSTS:
    hist_input(key, "  " + label, val, note)
HI["cost_check"] = r
put(ws, f"B{r}", "  Check: estimated split less reported expenses (should be 0)", italic=True)
put(ws, f"C{r}", f"=SUM(C{HI['cogs']}:C{HI['da']})-C{HI['expenses']}", fmt=M1, italic=True, fill=F_CHECK); r += 1
HI["pbt"] = r
put(ws, f"B{r}", "Profit before tax (derived)"); put(ws, f"C{r}", f"=C{HI['total_rev']}-C{HI['expenses']}", fmt=M1); r += 1
hist_input("npat", "NPAT (reported)", FY26["npat"][0], FY26["npat"][1])
HI["tax"] = r
put(ws, f"B{r}", "Income tax (derived)"); put(ws, f"C{r}", f"=C{HI['pbt']}-C{HI['npat']}", fmt=M1); r += 1
hist_input("cash", "Cash 30-Jun-2026", FY26["cash"][0], FY26["cash"][1])
hist_input("net_assets", "Net assets 30-Jun-2026", FY26["net_assets"][0], FY26["net_assets"][1])
hist_input("nwc", "Net working capital & other operating items (est.)", FY26["nwc"][0], FY26["nwc"][1])
hist_input("ocf", "Operating cash flow FY26A (memo)", FY26["ocf"][0], FY26["ocf"][1])
set_rows(ws)

# ========================================================================================
# OPCASE
# ========================================================================================
ws = S["OPCASE"]
title(ws, "Operating cases (OPCASE)", "Running block pulls the active case via OFFSET(case-1 cell, (OPCASE-1) x OpStep). "
      "FY27E = company status-quo plan (pre-close); FY28E-FY32E = sponsor plan.", "N")
widths(ws, {"A": 2, "B": 50, "C": 8, "D": 70, "E": 2, "F": 2, "G": 2, "H": 10, "I": 10, "J": 10, "K": 10, "L": 10, "M": 10, "N": 2})
put(ws, "B5", "Running operating case (OPCASE)", bold=True)
put(ws, "C5", "=MCASE!D8", fmt=INTF, bold=True, border=BOX, fill=F_RUN)
name("OPCASE", "OPCASE!$C$5")
BLOCK0, NDRV = 30, len(OP_DRIVERS)
OPSTEP = NDRV + 3
put(ws, "B6", "OpStep (rows between case blocks)")
put(ws, "C6", f"=ROW(B{BLOCK0 + OPSTEP})-ROW(B{BLOCK0})", fmt=INTF)
bar(ws, 8, "Running operating case (output - do not edit)", "B", "M", fill=F_OUT, color=WHITE)
put(ws, "B9", "Driver", bold=True); put(ws, "C9", "Unit", bold=True); put(ws, "D9", "Case description", bold=True)
for c, y in zip(H_TO_M, YEARS[1:]):
    put(ws, f"{c}9", y, bold=True, align="right")
OPRUN = {}
put(ws, "B10", "Case name / description", bold=True, fill=F_RUN)
put(ws, "C10", f"=OFFSET(C${BLOCK0},(OPCASE-1)*$C$6,0,1,1)", fill=F_RUN)
put(ws, "D10", f"=OFFSET(D${BLOCK0},(OPCASE-1)*$C$6,0,1,1)", fill=F_RUN, italic=True)
for i, (k, label, unit, fmt) in enumerate(OP_DRIVERS):
    rr = 11 + i
    OPRUN[k] = rr
    put(ws, f"B{rr}", label, fill=F_RUN)
    put(ws, f"C{rr}", unit, fill=F_RUN, color=GRAY)
    for c in H_TO_M:
        put(ws, f"{c}{rr}", f"=OFFSET({c}${BLOCK0 + 1 + i},(OPCASE-1)*$C$6,0,1,1)", fmt=fmt, fill=F_RUN)
for idx, (n, nm, desc, d) in enumerate(OPCASES):
    h = BLOCK0 + idx * OPSTEP
    bar(ws, h - 1, f"Operating case {n}: {nm}", "B", "M")
    put(ws, f"B{h}", f"Case {n}", bold=True)
    put(ws, f"C{h}", nm, bold=True, color=BLUE)
    put(ws, f"D{h}", desc, italic=True, color=BLUE, size=8)
    for i, (k, label, unit, fmt) in enumerate(OP_DRIVERS):
        rr = h + 1 + i
        put(ws, f"B{rr}", label)
        put(ws, f"C{rr}", unit, color=GRAY)
        for c, v in zip(H_TO_M, d[k]):
            put(ws, f"{c}{rr}", v, fmt=fmt)
put(ws, f"B{BLOCK0 + len(OPCASES) * OPSTEP}", "Notes: US/EU growth paths are analyst scenarios built around (i) LEO Pharma's dersimelagon NDA "
    "(priority review, FDA decision expected by end-Feb-2027), (ii) Disc Medicine's bitopertin (CRL 13-Feb-2026; APOLLO "
    "Phase 3 topline Q4 2026; decision mid-2027), (iii) EU year-round dosing (Sep-2025) and Canada approval (Jul-2026). "
    "FY27E costs follow the company's FY26-28 budget (avg A$55-58m p.a. ex-CBM, per PCR note Mar-2026).",
    italic=True, color=GRAY, size=8)
set_rows(ws)

# ========================================================================================
# FINCASE
# ========================================================================================
ws = S["FINCASE"]
title(ws, "Financing cases (FINCASE)", "Running output = OFFSET(case-1 cell, (FINCASE-1) x FinStep). Terms are indicative "
      "specialist private-credit terms for a single-product orphan asset (no term sheet).", "G")
widths(ws, {"A": 2, "B": 50, "C": 2, "D": 2, "E": 12, "F": 2, "G": 60})
put(ws, "B5", "Running financing case (FINCASE)", bold=True)
put(ws, "E5", "=MCASE!E8", fmt=INTF, bold=True, border=BOX, fill=F_RUN)
name("FINCASE", "FINCASE!$E$5")
FB0, NF = 34, len(FIN_DRIVERS)
FINSTEP = NF + 3
put(ws, "B6", "FinStep (rows between case blocks)")
put(ws, "E6", f"=ROW(B{FB0 + FINSTEP})-ROW(B{FB0})", fmt=INTF)
bar(ws, 8, "Running financing assumptions (output - do not edit)", "B", "G", fill=F_OUT, color=WHITE)
put(ws, "B9", "Case name", fill=F_RUN, bold=True)
put(ws, "E9", f"=OFFSET(E${FB0},(FINCASE-1)*$E$6,0,1,1)", fill=F_RUN)
FRUN = {}
for i, (k, label, fmt) in enumerate(FIN_DRIVERS):
    rr = 10 + i
    FRUN[k] = rr
    put(ws, f"B{rr}", label, fill=F_RUN)
    put(ws, f"E{rr}", f"=OFFSET(E${FB0 + 1 + i},(FINCASE-1)*$E$6,0,1,1)", fmt=fmt, fill=F_RUN)
for idx, (n, nm, vals) in enumerate(FINCASES):
    h = FB0 + idx * FINSTEP
    bar(ws, h - 1, f"Financing case {n}: {nm}", "B", "G")
    put(ws, f"B{h}", f"Case {n}", bold=True)
    put(ws, f"E{h}", nm, bold=True, color=BLUE)
    for i, ((k, label, fmt), v) in enumerate(zip(FIN_DRIVERS, vals)):
        put(ws, f"B{h + 1 + i}", label)
        put(ws, f"E{h + 1 + i}", v, fmt=fmt)
put(ws, "G10", "Leverage sized on FY27E (status-quo) Adj. EBITDA, i.e. before sponsor cost-out; "
    "5% amortisation reflects lender caution on a single product facing new entrants.", italic=True, color=GRAY, size=8, wrap=True)
set_rows(ws)

# ========================================================================================
# MCASE
# ========================================================================================
ws = S["MCASE"]
title(ws, "Master case driver (MCASE)", "Change C5 to run any case. Every tab keys off the running row (row 8).", "K")
widths(ws, {"A": 2, "B": 8, "C": 42, "D": 8, "E": 8, "F": 12, "G": 12, "H": 12, "I": 11, "J": 11, "K": 50})
put(ws, "B5", "MCASE", bold=True)
put(ws, "C5", DEFAULT_CASE, fmt=INTF, bold=True, border=BOX, fill=F_KEY)
name("MCASE", "MCASE!$C$5")
heads = ["Case #", "Case label", "Op case", "Fin case", "Offer price (A$/sh)", "Trans. exp. (% equity)",
         "Mgmt rollover (% equity)", "Exit multiple (x EBITDA)", "Min. cash (A$m)", "Comment"]
for i, h in enumerate(heads):
    put(ws, f"{get_column_letter(2 + i)}7", h, bold=True, color=WHITE, fill=F_OUT, wrap=True, align="center")
ws.row_dimensions[7].height = 36
put(ws, "B8", "=C5", fmt=INTF, bold=True, fill=F_RUN)
fmts = [None, INTF, INTF, PS, PCT2, PCT, MULT, M1]
for i, fmt in enumerate(fmts):
    col = get_column_letter(3 + i)
    put(ws, f"{col}8", f"=OFFSET({col}$9,$C$5-1,0,1,1)", fmt=fmt, bold=True, fill=F_RUN)
for n, lab, op, fin, px, te, ro, xm, mc in MCASES:
    rr = 8 + n
    put(ws, f"B{rr}", n, fmt=INTF, color=BLACK)
    put(ws, f"C{rr}", lab, color=BLUE)
    for col, v, fmt in zip("DEFGHIJ", [op, fin, px, te, ro, xm, mc], [INTF, INTF, PS, PCT2, PCT, MULT, M1]):
        put(ws, f"{col}{rr}", v, fmt=fmt)
put(ws, "K9", "Cases 1-20 = 5 operating x 4 financing cases at the base offer", italic=True, color=GRAY, size=8)
put(ws, "K29", "CEO rolls his c.6.8% stake (separate scheme class / GN19 protocol)", italic=True, color=GRAY, size=8)
put(ws, "K30", "c.+15% premium to undisturbed", italic=True, color=GRAY, size=8)
put(ws, "K31", "c.+52% premium; ~Mar-2026 trading levels", italic=True, color=GRAY, size=8)
put(ws, "K32", "Tests how much of the thesis is cost-out: management plan, no savings", italic=True, color=GRAY, size=8)
put(ws, "B35", "Offer price context: undisturbed A$8.24 (1-Sep-2026); A$10.75 = c.+30% premium; "
    "cash per share at close c.A$5.4. Exit multiples scale with the durability of each operating case.",
    italic=True, color=GRAY, size=8)
set_rows(ws)
ws.row_dimensions[7].height = 36

# ========================================================================================
# LBO
# ========================================================================================
ws = S["LBO"]
widths(ws, {"A": 2, "B": 52, "C": 44, "D": 11, "E": 7, "F": 2, "G": 10, "H": 10, "I": 10, "J": 10, "K": 10,
            "L": 10, "M": 10, "N": 10, "O": 2})
title(ws, "Project Daylight - Clinuvel Pharmaceuticals take-private: screening LBO",
      "A$ millions unless stated | FYE 30 June | Close assumed 30-Jun-2027 (end FY27E) | 5-year hold to FY32E | "
      "blue = input, black = formula, green = link", "O")
R = {}
row = [5]


def nr():
    return row[0]


def adv(n=1):
    row[0] += n


def section(text):
    adv(1)
    bar(ws, nr(), text, "B", "O")
    adv(1)


def line(key, label, vals=None, fmt=M1, note=None, bold=False, italic=False, top=False, fill=None, d=None, dfmt=None):
    rr = nr()
    R[key] = rr
    put(ws, f"B{rr}", label, bold=bold, italic=italic, fill=fill)
    if note:
        put(ws, f"C{rr}", note, italic=True, color=GRAY, size=8)
    if d is not None:
        put(ws, f"D{rr}", d, fmt=dfmt or fmt, bold=bold, fill=fill)
    for col, v in (vals or {}).items():
        put(ws, f"{col}{rr}", v, fmt=fmt, bold=bold, italic=italic, fill=fill, border=TOPLINE if top else None)
    adv(1)
    return rr


def yrs(fn, cols):
    return {c: fn(c) for c in cols}


# ---- control panel ----
put(ws, "B5", "CIRC switch (1 = average-balance interest; iterative calc on)", bold=True)
put(ws, "D5", 1, fmt=INTF, bold=True, border=BOX, fill=F_KEY)
name("CIRC", "LBO!$D$5")
put(ws, "B6", "Running master case (MCASE)"); put(ws, "D6", "=MCASE", fmt=INTF, color=GREEN, bold=True)
put(ws, "B7", "Case label"); put(ws, "D7", "=MCASE!C8")
put(ws, "B8", "Operating case / financing case"); put(ws, "D8", "=MCASE!D8", fmt=INTF); put(ws, "E8", "=MCASE!E8", fmt=INTF)
put(ws, "B9", "Operating case description"); put(ws, "D9", "=OPCASE!D10", italic=True, size=8)
row[0] = 10

section("Timeline")
line("fy", "Fiscal year", {c: y for c, y in zip(YC, YEARS)}, fmt=None, bold=True)
for c in YC:
    ws[f"{c}{R['fy']}"].alignment = Alignment(horizontal="right")
line("date", "Period end", {"G": "=DATE(2026,6,30)", **yrs(lambda c: f"=DATE(YEAR({PREV[c]}{nr()})+1,6,30)", YC[1:])}, fmt=DATEF)
ws[f"G{R['date']}"].font = Font(name=FONT, size=9, color=BLUE)
line("period", "Period", {"G": "Actual", "H": "Status quo", **{c: "Sponsor" for c in PROJ}}, fmt=None, italic=True)
for c in YC:
    ws[f"{c}{R['period']}"].alignment = Alignment(horizontal="right")
line("years", "Years since close", yrs(lambda c: f"=YEAR({c}{R['date']})-YEAR($H${R['date']})", PROJ), fmt=INTF)

# ---- transaction assumptions ----
section("Transaction assumptions")
T = TXN
line("undisturbed", "Undisturbed share price (A$)", d=T["undisturbed"][0], fmt=PS, note=T["undisturbed"][1])
line("offer", "Offer price per share (A$)", d="=MCASE!F8", fmt=PS, note="MCASE running row", bold=True)
line("premium", "Premium to undisturbed", d=f"=D{R['offer']}/D{R['undisturbed']}-1", fmt=PCT)
line("shares_basic", "Basic shares outstanding (m)", d=T["shares_basic"][0], fmt=SH, note=T["shares_basic"][1])
line("dilutive", "Dilutive securities (m)", d=T["dilutive"][0], fmt=SH, note=T["dilutive"][1])
line("shares_fd", "Fully diluted shares (m)", d=f"=D{R['shares_basic']}+D{R['dilutive']}", fmt=SH)
line("eqpp", "Equity purchase price", d=f"=D{R['offer']}*D{R['shares_fd']}", bold=True)
line("cash_pre", "Clinuvel cash at close (pre-deal, FY27E)", d="=H{cf_end}", note="Status-quo FY27E closing cash (cash flow section)")
line("debt_existing", "Existing debt refinanced", d=T["debt_existing"][0], note=T["debt_existing"][1])
line("min_cash", "Minimum operating cash retained", d="=MCASE!J8", note="Named MIN_CASH")
name("MIN_CASH", f"LBO!$D${R['min_cash']}")
line("cash_applied", "Clinuvel cash applied to the transaction", d=f"=MAX(0,D{R['cash_pre']}-D{R['min_cash']})")
line("tev", "Implied transaction enterprise value (TEV)", d=f"=D{R['eqpp']}+D{R['debt_existing']}-D{R['cash_pre']}", bold=True)
line("ebitda_ltm", "LTM Adj. EBITDA at close (FY27E, status quo)", d="=H{adj_ebitda}")
line("ebitda_fy26", "FY26A Adj. EBITDA", d="=G{adj_ebitda}")
line("ebitda_fy28", "FY28E Adj. EBITDA (first sponsor year)", d="=I{adj_ebitda}")
line("tev_ltm_x", "TEV / LTM (FY27E) Adj. EBITDA", d=f"=D{R['tev']}/D{R['ebitda_ltm']}", fmt=MULT)
line("tev_fy26_x", "TEV / FY26A Adj. EBITDA", d=f"=D{R['tev']}/D{R['ebitda_fy26']}", fmt=MULT)
line("tev_fy28_x", "TEV / FY28E Adj. EBITDA (post cost-out)", d=f"=D{R['tev']}/D{R['ebitda_fy28']}", fmt=MULT)
line("tev_rev_x", "TEV / FY27E revenue", d="=D{tev}/H{rev_total}", fmt=MULT)
line("pe_x", "Equity purchase price / FY26A NPAT", d="=D{eqpp}/G{ni}", fmt=MULT)
line("cash_ps", "Cash per fully diluted share at close (A$)", d=f"=D{R['cash_pre']}/D{R['shares_fd']}", fmt=PS)
line("trans_pct", "Transaction expenses (% of equity purchase price)", d="=MCASE!G8", fmt=PCT2)
line("trans_exp", "Transaction expenses", d=f"=D{R['trans_pct']}*D{R['eqpp']}", note="Advisers, legal, IER, scheme, FIRB/ACCC, US counsel")
line("roll_pct", "Management rollover (% of equity purchase price)", d="=MCASE!H8", fmt=PCT)
line("rollover", "Management rollover", d=f"=D{R['roll_pct']}*D{R['eqpp']}")
line("exit_mult", "Exit multiple (x LTM Adj. EBITDA)", d="=MCASE!I8", fmt=MULT)
for k, lab, fmt in [("exit_cost_pct", "Exit costs (% of TEV)", PCT), ("tax_rate", "Tax rate", PCT),
                    ("deposit_rate", "Interest yield on cash", PCT2), ("dps_fy27", "Dividend paid in FY27E (A$/sh)", PS),
                    ("da_run", "D&A run-rate", M1), ("us_incr", "Incremental US HQ / Nasdaq run-rate cost", M1),
                    ("dist_pct", "Excess-cash distribution to equity (once debt repaid)", PCT),
                    ("min_debt_incr", "Debt sizing increment (A$m)", M1)]:
    line(k, lab, d=T[k][0], fmt=fmt, note=T[k][1])

# ---- financing assumptions ----
section("Financing assumptions (from FINCASE)")
hdr = nr()
for col, h in zip("GHIJKLMN", ["Lev. (x)", "A$m", "Margin/cpn", "Floor", "Fee %", "Fee A$m", "Tenor", "Amort."]):
    put(ws, f"{col}{hdr}", h, bold=True, align="right")
put(ws, f"B{hdr}", "Tranche", bold=True)
adv(1)
F = FRUN
fr = nr(); R["rcf_fin"] = fr
put(ws, f"B{fr}", "Revolving credit facility (commitment; undrawn at close)")
put(ws, f"H{fr}", f"=FINCASE!E{F['rcf_commit']}", fmt=M1)
put(ws, f"I{fr}", f"=FINCASE!E{F['rcf_margin']}", fmt=PCT2)
put(ws, f"K{fr}", f"=FINCASE!E{F['rcf_fee']}", fmt=PCT2)
put(ws, f"L{fr}", f"=H{fr}*K{fr}", fmt=M1)
put(ws, f"M{fr}", f"=FINCASE!E{F['rcf_tenor']}", fmt=INTF)
put(ws, f"N{fr}", 0, fmt=PCT)
adv(1)
fr = nr(); R["tl_fin"] = fr
put(ws, f"B{fr}", "Unitranche term loan")
put(ws, f"G{fr}", f"=FINCASE!E{F['tl_lev']}", fmt=MULT)
put(ws, f"H{fr}", f"=MROUND(G{fr}*$D${R['ebitda_ltm']},$D${R['min_debt_incr']})", fmt=M1)
put(ws, f"I{fr}", f"=FINCASE!E{F['tl_margin']}", fmt=PCT2)
put(ws, f"J{fr}", f"=FINCASE!E{F['tl_floor']}", fmt=PCT2)
put(ws, f"K{fr}", f"=FINCASE!E{F['tl_fee']}", fmt=PCT2)
put(ws, f"L{fr}", f"=H{fr}*K{fr}", fmt=M1)
put(ws, f"M{fr}", f"=FINCASE!E{F['tl_tenor']}", fmt=INTF)
put(ws, f"N{fr}", f"=FINCASE!E{F['tl_amort']}", fmt=PCT)
adv(1)
fr = nr(); R["pik_fin"] = fr
put(ws, f"B{fr}", "HoldCo PIK notes")
put(ws, f"G{fr}", f"=FINCASE!E{F['pik_lev']}", fmt=MULT)
put(ws, f"H{fr}", f"=MROUND(G{fr}*$D${R['ebitda_ltm']},$D${R['min_debt_incr']})", fmt=M1)
put(ws, f"I{fr}", f"=FINCASE!E{F['pik_rate']}", fmt=PCT2)
put(ws, f"K{fr}", f"=FINCASE!E{F['pik_fee']}", fmt=PCT2)
put(ws, f"L{fr}", f"=H{fr}*K{fr}", fmt=M1)
put(ws, f"M{fr}", f"=FINCASE!E{F['pik_tenor']}", fmt=INTF)
put(ws, f"N{fr}", 0, fmt=PCT)
adv(1)
fr = nr(); R["fin_total"] = fr
put(ws, f"B{fr}", "Total funded debt / total fees", bold=True)
put(ws, f"G{fr}", f"=G{R['tl_fin']}+G{R['pik_fin']}", fmt=MULT, bold=True, border=TOPLINE)
put(ws, f"H{fr}", f"=H{R['tl_fin']}+H{R['pik_fin']}", fmt=M1, bold=True, border=TOPLINE)
put(ws, f"L{fr}", f"=SUM(L{R['rcf_fin']}:L{R['pik_fin']})", fmt=M1, bold=True, border=TOPLINE)
adv(1)
line("rcf_cfee", "Revolver commitment fee (on undrawn)", {"I": f"=FINCASE!E{F['rcf_cfee']}"}, fmt=PCT2)
line("rcf_draw", "Revolver drawn at close", {"H": 0.0}, fmt=M1)
line("cov_lev", "Covenant: max net debt / EBITDA", d=f"=FINCASE!E{F['cov_lev']}", fmt=MULT)
line("cov_icr", "Covenant: min EBITDA / cash interest", d=f"=FINCASE!E{F['cov_icr']}", fmt=MULT)

# ---- sources & uses ----
section("Sources & uses at close (30-Jun-2027)")
put(ws, f"D{nr()}", "A$m", bold=True, align="right"); put(ws, f"E{nr()}", "%", bold=True, align="right")
put(ws, f"B{nr()}", "Uses", bold=True); adv(1)
line("u_eq", "Purchase of equity (fully diluted)", d=f"=D{R['eqpp']}")
line("u_refi", "Refinance existing debt", d=f"=D{R['debt_existing']}")
line("u_trans", "Transaction expenses", d=f"=D{R['trans_exp']}")
line("u_fin", "Financing fees", d=f"=L{R['fin_total']}")
line("u_total", "Total uses", d=f"=SUM(D{R['u_eq']}:D{R['u_fin']})", bold=True)
put(ws, f"B{nr()}", "Sources", bold=True); adv(1)
line("s_cash", "Clinuvel balance-sheet cash applied", d=f"=D{R['cash_applied']}")
line("s_rcf", "Revolver drawn", d=f"=H{R['rcf_draw']}")
line("s_tl", "Unitranche term loan", d=f"=H{R['tl_fin']}")
line("s_pik", "HoldCo PIK notes", d=f"=H{R['pik_fin']}")
line("s_roll", "Management rollover", d=f"=D{R['rollover']}")
line("s_spon", "Sponsor equity (plug)", d=f"=D{R['u_total']}-SUM(D{R['s_cash']}:D{R['s_roll']})", bold=True)
line("s_total", "Total sources", d=f"=SUM(D{R['s_cash']}:D{R['s_spon']})", bold=True)
for k in ["u_eq", "u_refi", "u_trans", "u_fin", "u_total"]:
    put(ws, f"E{R[k]}", f"=D{R[k]}/$D${R['u_total']}", fmt=PCT)
for k in ["s_cash", "s_rcf", "s_tl", "s_pik", "s_roll", "s_spon", "s_total"]:
    put(ws, f"E{R[k]}", f"=D{R[k]}/$D${R['s_total']}", fmt=PCT)
line("su_check", "Check: sources less uses (should be 0)", d=f"=D{R['s_total']}-D{R['u_total']}", italic=True, fill=F_CHECK)
line("eq_total", "Total equity invested (sponsor + rollover)", d=f"=D{R['s_spon']}+D{R['s_roll']}", bold=True)
line("eq_pct_new", "Equity / (equity + new debt)", d=f"=D{R['eq_total']}/(D{R['eq_total']}+D{R['s_rcf']}+D{R['s_tl']}+D{R['s_pik']})", fmt=PCT)
line("eq_pct_total", "Equity / total sources", d=f"=D{R['eq_total']}/D{R['s_total']}", fmt=PCT)
line("pf_nlev", "PF net debt / LTM Adj. EBITDA", d=f"=(D{R['s_rcf']}+D{R['s_tl']}+D{R['s_pik']}-D{R['min_cash']})/D{R['ebitda_ltm']}", fmt=MULT)
line("non_price_inv", "Memo: equity invested excl. price-driven items (fees - cash applied - new debt)",
     d=f"=D{R['u_fin']}-D{R['s_cash']}-D{R['s_rcf']}-D{R['s_tl']}-D{R['s_pik']}", note="Invested(p) = p x FD shares x (1+fee%) + this")

# ---- purchase accounting ----
section("Purchase accounting (simplified: excess over book to goodwill & intangibles)")
line("pa_eqpp", "Equity purchase price", d=f"=D{R['eqpp']}")
line("pa_book", "Less: pre-deal book equity (FY27E)", d="=-G{pf_equity}")
line("pa_gw", "Goodwill & intangibles created (plug)", d=f"=D{R['pa_eqpp']}+D{R['pa_book']}", bold=True)

# ---- PF closing balance sheet ----
section("Pro forma closing balance sheet (30-Jun-2027)")
ph = nr()
for col, h in zip("GHI", ["Pre-deal", "Adjust.", "PF close"]):
    put(ws, f"{col}{ph}", h, bold=True, align="right")
adv(1)
line("pf_cash", "Cash", {"G": "=H{cf_end}", "H": f"=-D{R['cash_applied']}+D{R['s_rcf']}", "I": "=G{r}+H{r}"})
line("pf_nwc", "Net working capital & other", {"G": "=H{bs_nwc}", "H": 0, "I": "=G{r}+H{r}"})
line("pf_ppe", "PP&E & other non-current", {"G": "=H{bs_ppe}", "H": 0, "I": "=G{r}+H{r}"})
line("pf_gw", "Goodwill & intangibles (new)", {"G": 0, "H": f"=D{R['pa_gw']}", "I": "=G{r}+H{r}"})
line("pf_ff", "Capitalised financing fees", {"G": 0, "H": f"=D{R['u_fin']}", "I": "=G{r}+H{r}"})
line("pf_ta", "Total assets", yrs(lambda c: f"=SUM({c}{R['pf_cash']}:{c}{R['pf_ff']})", "GHI"), bold=True, top=True)
line("pf_rcf", "Revolver", {"G": 0, "H": f"=D{R['s_rcf']}", "I": "=G{r}+H{r}"})
line("pf_tl", "Unitranche term loan", {"G": 0, "H": f"=D{R['s_tl']}", "I": "=G{r}+H{r}"})
line("pf_pik", "HoldCo PIK notes", {"G": 0, "H": f"=D{R['s_pik']}", "I": "=G{r}+H{r}"})
line("pf_debt", "Total debt", yrs(lambda c: f"=SUM({c}{R['pf_rcf']}:{c}{R['pf_pik']})", "GHI"), bold=True, top=True)
line("pf_equity", "Shareholders' equity", {"G": "=G{bs_equity}+H{ni}+H{cf_div}",
                                           "H": f"=-G{{r}}+D{R['eq_total']}-D{R['u_trans']}", "I": "=G{r}+H{r}"})
line("pf_tle", "Total liabilities & equity", yrs(lambda c: f"={c}{R['pf_debt']}+{c}{R['pf_equity']}", "GHI"), bold=True, top=True)
line("pf_check", "Check: assets less L&E (should be 0)", yrs(lambda c: f"={c}{R['pf_ta']}-{c}{R['pf_tle']}", "GHI"), italic=True, fill=F_CHECK)

# ---- income statement ----
section("Income statement")
yh = nr()
for c, y in zip(YC, YEARS):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
O = OPRUN
HS = "Hist!C"
line("rev_us", "SCENESSE revenue - United States", {"G": f"={HS}{HI['us_rev']}", **yrs(lambda c: f"={PREV[c]}{{r}}*(1+OPCASE!{c}{O['us_g']})", H_TO_M)})
line("rev_eu", "SCENESSE revenue - Europe & rest of world", {"G": f"={HS}{HI['eu_rev']}", **yrs(lambda c: f"={PREV[c]}{{r}}*(1+OPCASE!{c}{O['eu_g']})", H_TO_M)})
line("rev_vit", "Vitiligo revenue", {"G": 0, **yrs(lambda c: f"=OPCASE!{c}{O['vit_rev']}", H_TO_M)})
line("rev_oth", "Other new products", {"G": 0, **yrs(lambda c: f"=OPCASE!{c}{O['oth_rev']}", H_TO_M)})
line("rev_total", "Total revenue", yrs(lambda c: f"=SUM({c}{R['rev_us']}:{c}{R['rev_oth']})", YC), bold=True, top=True)
line("rev_g", "  growth", yrs(lambda c: f"={c}{R['rev_total']}/{PREV[c]}{R['rev_total']}-1", H_TO_M), fmt=PCT, italic=True)
line("rev_us_share", "  US share of revenue", yrs(lambda c: f"={c}{R['rev_us']}/{c}{R['rev_total']}", YC), fmt=PCT, italic=True)
line("cogs", "Cost of goods, supply & distribution", {"G": f"=-{HS}{HI['cogs']}", **yrs(lambda c: f"=-{c}{R['rev_total']}*OPCASE!{c}{O['cogs']}", H_TO_M)})
line("gp", "Gross profit", yrs(lambda c: f"={c}{R['rev_total']}+{c}{R['cogs']}", YC), bold=True, top=True)
line("gm", "  gross margin", yrs(lambda c: f"={c}{R['gp']}/{c}{R['rev_total']}", YC), fmt=PCT, italic=True)
line("opex_core", "Core SCENESSE commercial, medical, PV & regulatory", {"G": f"=-{HS}{HI['core']}", **yrs(lambda c: f"={PREV[c]}{{r}}*(1+OPCASE!{c}{O['infl']})", H_TO_M)})
line("opex_def", "Competitive-defence spend", {"G": 0, **yrs(lambda c: f"=OPCASE!{c}{O['def']}", H_TO_M)})
line("opex_corp", "Corporate, listing & G&A (status-quo run-rate)", {"G": f"=-{HS}{HI['corp']}", **yrs(lambda c: f"={PREV[c]}{{r}}*(1+OPCASE!{c}{O['infl']})", H_TO_M)})
line("opex_us", "Incremental US HQ / Nasdaq / SEC costs", {"G": 0, "H": f"=-$D${R['us_incr']}", **yrs(lambda c: f"={PREV[c]}{{r}}*(1+OPCASE!{c}{O['infl']})", PROJ)})
line("opex_sav", "Sponsor corporate cost savings", {"G": 0, **yrs(lambda c: f"=-({c}{R['opex_corp']}+{c}{R['opex_us']})*OPCASE!{c}{O['sav']}", H_TO_M)})
line("opex_cbm", "Communications, branding, marketing & consumer", {"G": f"=-{HS}{HI['cbm']}", **yrs(lambda c: f"=OPCASE!{c}{O['cbm']}", H_TO_M)})
line("opex_vit", "Vitiligo clinical program", {"G": f"=-{HS}{HI['vit']}", **yrs(lambda c: f"=OPCASE!{c}{O['vit']}", H_TO_M)})
line("opex_rd", "Other R&D (Singapore, PRENUMBRA, NEURACTHEL, XP/VP)", {"G": f"=-{HS}{HI['rd']}", **yrs(lambda c: f"=OPCASE!{c}{O['rd']}", H_TO_M)})
line("opex_vitl", "Vitiligo launch & commercial", {"G": 0, **yrs(lambda c: f"=OPCASE!{c}{O['vitl']}", H_TO_M)})
line("opex_total", "Total operating expenses (ex-D&A)", yrs(lambda c: f"=SUM({c}{R['opex_core']}:{c}{R['opex_vitl']})", YC), bold=True, top=True)
line("adj_ebitda", "Adjusted EBITDA", yrs(lambda c: f"={c}{R['gp']}+{c}{R['opex_total']}", YC), bold=True, fill=F_KEY)
line("ebitda_m", "  Adj. EBITDA margin", yrs(lambda c: f"={c}{R['adj_ebitda']}/{c}{R['rev_total']}", YC), fmt=PCT, italic=True)
line("disc_spend", "  memo: discretionary pipeline & CBM spend", yrs(lambda c: f"={c}{R['opex_cbm']}+{c}{R['opex_vit']}+{c}{R['opex_rd']}+{c}{R['opex_vitl']}", YC), italic=True)
line("nonrec", "Non-recurring items", {"G": 0, **yrs(lambda c: f"=OPCASE!{c}{O['nonrec']}", H_TO_M)})
line("ebitda", "Reported EBITDA", yrs(lambda c: f"={c}{R['adj_ebitda']}+{c}{R['nonrec']}", YC), bold=True, top=True)
line("da", "Depreciation & amortisation", {"G": f"=-{HS}{HI['da']}", **yrs(lambda c: f"=-$D${R['da_run']}", H_TO_M)})
line("ebit", "EBIT", yrs(lambda c: f"={c}{R['ebitda']}+{c}{R['da']}", YC), bold=True, top=True)
line("int_inc", "Interest & other income", {"G": f"={HS}{HI['other_inc']}", **yrs(lambda c: f"=$D${R['deposit_rate']}*IF(CIRC=1,AVERAGE({c}{{cf_beg}},{c}{{cf_end}}),{c}{{cf_beg}})", H_TO_M)})
line("int_cash", "Cash interest expense", {"G": 0, "H": 0, **yrs(lambda c: f"=IF(CIRC=1,-{c}{{int_total_cash}},0)", PROJ)})
line("int_noncash", "Non-cash interest (PIK + fee amortisation)", {"G": 0, "H": 0, **yrs(lambda c: f"=IF(CIRC=1,-{c}{{int_noncash}},0)", PROJ)})
line("ebt", "Profit before tax", yrs(lambda c: f"={c}{R['ebit']}+{c}{R['int_inc']}+{c}{R['int_cash']}+{c}{R['int_noncash']}", YC), bold=True, top=True)
line("tax", "Income tax", {"G": f"=-{HS}{HI['tax']}", **yrs(lambda c: f"=-MAX(0,{c}{R['ebt']}*$D${R['tax_rate']})", H_TO_M)})
line("ni", "Net income", yrs(lambda c: f"={c}{R['ebt']}+{c}{R['tax']}", YC), bold=True, top=True)
line("ni_check", "  check: FY26A NPAT vs reported (should be 0)", {"G": f"=G{{ni}}-{HS}{HI['npat']}"}, italic=True, fill=F_CHECK)

# ---- balance sheet ----
section("Balance sheet (FY27E column = pro forma post-close)")
yh = nr()
for c, y in zip(YC, ["FY26A", "PF FY27E"] + YEARS[2:]):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
line("bs_cash", "Cash", {"G": f"={HS}{HI['cash']}", "H": "=I{pf_cash}", **yrs(lambda c: f"={c}{{cf_end}}", PROJ)})
line("bs_nwc", "Net working capital & other", {"G": f"={HS}{HI['nwc']}", **yrs(lambda c: f"={c}{R['rev_total']}*OPCASE!{c}{O['nwc']}", H_TO_M)})
line("bs_ppe", "PP&E & other non-current (derived FY26A)", {"G": f"={HS}{HI['net_assets']}-G{R['bs_cash']}-G{R['bs_nwc']}", **yrs(lambda c: f"={PREV[c]}{{r}}-{c}{{cf_capex}}+{c}{R['da']}", H_TO_M)})
line("bs_gw", "Goodwill & intangibles", {"G": 0, "H": "=I{pf_gw}", **yrs(lambda c: f"={PREV[c]}{{r}}", PROJ)})
line("bs_ff", "Capitalised financing fees", {"G": 0, "H": "=I{pf_ff}", **yrs(lambda c: f"={PREV[c]}{{r}}-{c}{{fa_total}}", PROJ)})
line("bs_ta", "Total assets", yrs(lambda c: f"=SUM({c}{R['bs_cash']}:{c}{R['bs_ff']})", YC), bold=True, top=True)
line("bs_rcf", "Revolver", {"G": 0, "H": "=I{pf_rcf}", **yrs(lambda c: f"={c}{{rcf_end}}", PROJ)})
line("bs_tl", "Unitranche term loan", {"G": 0, "H": "=I{pf_tl}", **yrs(lambda c: f"={c}{{tl_end}}", PROJ)})
line("bs_pik", "HoldCo PIK notes", {"G": 0, "H": "=I{pf_pik}", **yrs(lambda c: f"={c}{{pik_end}}", PROJ)})
line("bs_debt", "Total debt", yrs(lambda c: f"=SUM({c}{R['bs_rcf']}:{c}{R['bs_pik']})", YC), bold=True, top=True)
line("bs_equity", "Shareholders' equity", {"G": f"={HS}{HI['net_assets']}", "H": "=I{pf_equity}", **yrs(lambda c: f"={PREV[c]}{{r}}+{c}{{ni}}+{c}{{cf_dist}}", PROJ)})
line("bs_tle", "Total liabilities & equity", yrs(lambda c: f"={c}{R['bs_debt']}+{c}{R['bs_equity']}", YC), bold=True, top=True)
line("bs_check", "Check: assets less L&E (should be 0)", yrs(lambda c: f"=ROUND({c}{R['bs_ta']}-{c}{R['bs_tle']},4)", YC), italic=True, fill=F_CHECK)

# ---- cash flow ----
section("Cash flow statement (FY27E = status quo pre-close; FY28E+ = sponsor ownership)")
yh = nr()
for c, y in zip(H_TO_M, YEARS[1:]):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
line("cf_ni", "Net income", yrs(lambda c: f"={c}{R['ni']}", H_TO_M))
line("cf_da", "Add: D&A", yrs(lambda c: f"=-{c}{R['da']}", H_TO_M))
line("cf_nci", "Add: non-cash interest", yrs(lambda c: f"=-{c}{R['int_noncash']}", H_TO_M))
line("cf_nwc", "(Increase) / decrease in NWC", yrs(lambda c: f"=-({c}{R['bs_nwc']}-{PREV[c]}{R['bs_nwc']})", H_TO_M))
line("cf_cfo", "Cash flow from operations", yrs(lambda c: f"=SUM({c}{R['cf_ni']}:{c}{R['cf_nwc']})", H_TO_M), bold=True, top=True)
line("cf_capex", "Capex", yrs(lambda c: f"=OPCASE!{c}{O['capex']}", H_TO_M))
line("cf_fcf", "Free cash flow (after interest & tax)", yrs(lambda c: f"={c}{R['cf_cfo']}+{c}{R['cf_capex']}", H_TO_M), bold=True, top=True)
line("cf_div", "Dividends paid (status quo, pre-close)", {"H": f"=-$D${R['dps_fy27']}*$D${R['shares_basic']}", **{c: 0 for c in PROJ}})
line("cf_mand", "Mandatory debt amortisation", {"H": 0, **yrs(lambda c: f"={c}{{ds_mand}}", PROJ)})
line("cf_opt", "Optional prepayment / (revolver draw)", {"H": 0, **yrs(lambda c: f"={c}{{ds_opt}}", PROJ)})
line("cf_dist", "Distributions to equity holders", {"H": 0, **yrs(lambda c: f"=-MAX(0,{c}{{cf_beg}}+{c}{R['cf_fcf']}+{c}{R['cf_div']}+{c}{R['cf_mand']}+{c}{R['cf_opt']}-MIN_CASH)*$D${R['dist_pct']}", PROJ)})
line("cf_net", "Net change in cash", yrs(lambda c: f"={c}{R['cf_fcf']}+{c}{R['cf_div']}+{c}{R['cf_mand']}+{c}{R['cf_opt']}+{c}{R['cf_dist']}", H_TO_M), bold=True, top=True)
line("cf_beg", "Beginning cash (FY28E = PF cash after close)", yrs(lambda c: f"={PREV[c]}{R['bs_cash']}", H_TO_M))
line("cf_end", "Ending cash", yrs(lambda c: f"={c}{R['cf_beg']}+{c}{R['cf_net']}", H_TO_M), bold=True)
line("cf_avail", "Memo: cash available for debt repayment", yrs(lambda c: f"={c}{R['cf_beg']}+{c}{R['cf_fcf']}+{c}{R['cf_div']}-MIN_CASH", PROJ), italic=True)
line("ufcf", "Memo: unlevered free cash flow", yrs(lambda c: f"={c}{R['ebitda']}+{c}{R['cf_nwc']}+{c}{R['cf_capex']}-MAX(0,{c}{R['ebit']})*$D${R['tax_rate']}", H_TO_M), italic=True)

# ---- debt schedule ----
section("Debt schedule")
yh = nr()
for c, y in zip(YC[1:], ["PF FY27E"] + YEARS[2:]):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
line("ds_avail", "Cash available for debt repayment", yrs(lambda c: f"={c}{R['cf_avail']}", PROJ))
line("ds_mand", "Total mandatory amortisation", yrs(lambda c: f"={c}{{rcf_mand}}+{c}{{tl_mand}}+{c}{{pik_mand}}", PROJ))
line("ds_opt_avail", "Cash available for optional prepayment", yrs(lambda c: f"={c}{R['ds_avail']}+{c}{R['ds_mand']}", PROJ), bold=True)
put(ws, f"B{nr()}", "Revolver", bold=True); adv(1)
line("rcf_mand", "  Mandatory amortisation", {c: 0 for c in PROJ})
line("rcf_opt", "  Optional prepayment / (draw)", yrs(lambda c: f"=IF({c}{R['ds_opt_avail']}>0,-MAX(0,MIN({PREV[c]}{{rcf_end}},{c}{R['ds_opt_avail']})),-MIN(0,{c}{R['ds_opt_avail']}))", PROJ))
line("rcf_end", "  Ending balance", {"H": "=I{pf_rcf}", **yrs(lambda c: f"=MAX(0,{PREV[c]}{{r}}+{c}{R['rcf_mand']}+{c}{R['rcf_opt']})", PROJ)}, bold=True)
put(ws, f"B{nr()}", "Unitranche term loan", bold=True); adv(1)
line("tl_mand", "  Mandatory amortisation", yrs(lambda c: f"=-MAX(0,MIN($H{{tl_end}}*$N${R['tl_fin']},{PREV[c]}{{tl_end}}))", PROJ))
line("tl_opt", "  Optional prepayment (cash sweep)", yrs(lambda c: f"=-MAX(0,MIN({PREV[c]}{{tl_end}}+{c}{R['tl_mand']},{c}{R['ds_opt_avail']}+{c}{R['rcf_opt']}))", PROJ))
line("tl_end", "  Ending balance", {"H": "=I{pf_tl}", **yrs(lambda c: f"=MAX(0,{PREV[c]}{{r}}+{c}{R['tl_mand']}+{c}{R['tl_opt']})", PROJ)}, bold=True)
put(ws, f"B{nr()}", "HoldCo PIK notes", bold=True); adv(1)
line("pik_mand", "  Mandatory amortisation", {c: 0 for c in PROJ})
line("pik_acc", "  PIK interest accrued", yrs(lambda c: f"={c}{{int_pik}}", PROJ))
line("pik_opt", "  Optional prepayment", yrs(lambda c: f"=-MAX(0,MIN({PREV[c]}{{pik_end}}+{c}{R['pik_acc']},{c}{R['ds_opt_avail']}+{c}{R['rcf_opt']}+{c}{R['tl_opt']}))", PROJ))
line("pik_end", "  Ending balance", {"H": "=I{pf_pik}", **yrs(lambda c: f"=MAX(0,{PREV[c]}{{r}}+{c}{R['pik_mand']}+{c}{R['pik_opt']}+{c}{R['pik_acc']})", PROJ)}, bold=True)
line("ds_opt", "Total optional prepayment / (draw)", yrs(lambda c: f"={c}{R['rcf_opt']}+{c}{R['tl_opt']}+{c}{R['pik_opt']}", PROJ), bold=True, top=True)
line("ds_total", "Total debt outstanding", {"H": f"=H{R['rcf_end']}+H{R['tl_end']}+H{R['pik_end']}", **yrs(lambda c: f"={c}{R['rcf_end']}+{c}{R['tl_end']}+{c}{R['pik_end']}", PROJ)}, bold=True)

# ---- interest ----
section("Interest expense schedule")
yh = nr()
for c, y in zip(PROJ, YEARS[2:]):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
line("base_rate", "Base rate (BBSY / SOFR proxy)", {c: v for c, v in zip(PROJ, BASE_RATE)}, fmt=PCT2, note="ASSUMPTION: flat curve")
line("rcf_avg", "Revolver: average drawn balance", yrs(lambda c: f"=IF(CIRC=1,AVERAGE({PREV[c]}{R['rcf_end']},{c}{R['rcf_end']}),{PREV[c]}{R['rcf_end']})", PROJ))
line("rcf_int", "Revolver: interest + commitment fee", yrs(lambda c: f"={c}{R['rcf_avg']}*({c}{R['base_rate']}+$I${R['rcf_fin']})+MAX(0,$H${R['rcf_fin']}-{c}{R['rcf_avg']})*$I${R['rcf_cfee']}", PROJ))
line("tl_avg", "Term loan: average balance", yrs(lambda c: f"=IF(CIRC=1,AVERAGE({PREV[c]}{R['tl_end']},{c}{R['tl_end']}),{PREV[c]}{R['tl_end']})", PROJ))
line("tl_rate", "Term loan: all-in rate", yrs(lambda c: f"=MAX($J${R['tl_fin']},{c}{R['base_rate']})+$I${R['tl_fin']}", PROJ), fmt=PCT2)
line("tl_int", "Term loan: cash interest", yrs(lambda c: f"={c}{R['tl_avg']}*{c}{R['tl_rate']}", PROJ))
line("pik_avg", "PIK notes: average balance", yrs(lambda c: f"=IF(CIRC=1,AVERAGE({PREV[c]}{R['pik_end']},{c}{R['pik_end']}),{PREV[c]}{R['pik_end']})", PROJ))
line("int_pik", "PIK notes: interest accrued (non-cash)", yrs(lambda c: f"={c}{R['pik_avg']}*$I${R['pik_fin']}", PROJ))
line("int_total_cash", "Total cash interest expense", yrs(lambda c: f"={c}{R['rcf_int']}+{c}{R['tl_int']}", PROJ), bold=True, top=True)
line("fa_rcf", "Fee amortisation - revolver", yrs(lambda c: f"=MAX(0,MIN($L${R['rcf_fin']}-SUM($H{{r}}:{PREV[c]}{{r}}),$L${R['rcf_fin']}/$M${R['rcf_fin']}))", PROJ))
line("fa_tl", "Fee amortisation - term loan", yrs(lambda c: f"=MAX(0,MIN($L${R['tl_fin']}-SUM($H{{r}}:{PREV[c]}{{r}}),$L${R['tl_fin']}/$M${R['tl_fin']}))", PROJ))
line("fa_pik", "Fee amortisation - PIK notes", yrs(lambda c: f"=MAX(0,MIN($L${R['pik_fin']}-SUM($H{{r}}:{PREV[c]}{{r}}),$L${R['pik_fin']}/$M${R['pik_fin']}))", PROJ))
line("fa_total", "Total financing fee amortisation", yrs(lambda c: f"=SUM({c}{R['fa_rcf']}:{c}{R['fa_pik']})", PROJ), top=True)
line("int_noncash", "Total non-cash interest", yrs(lambda c: f"={c}{R['int_pik']}+{c}{R['fa_total']}", PROJ), bold=True)
line("int_total", "Total interest expense", yrs(lambda c: f"={c}{R['int_total_cash']}+{c}{R['int_noncash']}", PROJ), bold=True)

# ---- credit metrics ----
section("Credit metrics")
yh = nr()
for c, y in zip(YC[1:], ["PF FY27E"] + YEARS[2:]):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
CM = ["H"] + PROJ
line("cm_debt", "Total debt", yrs(lambda c: f"={c}{R['bs_debt']}", CM))
line("cm_cash", "Cash", yrs(lambda c: f"={c}{R['bs_cash']}", CM))
line("cm_nd", "Net debt / (net cash)", yrs(lambda c: f"={c}{R['cm_debt']}-{c}{R['cm_cash']}", CM), bold=True)
line("cm_ebitda", "LTM Adj. EBITDA", yrs(lambda c: f"={c}{R['adj_ebitda']}", CM))
line("cm_lev", "Total debt / EBITDA", yrs(lambda c: f"={c}{R['cm_debt']}/{c}{R['cm_ebitda']}", CM), fmt=MULT2)
line("cm_nlev", "Net debt / EBITDA", yrs(lambda c: f"={c}{R['cm_nd']}/{c}{R['cm_ebitda']}", CM), fmt=MULT2)
line("cm_icr", "EBITDA / cash interest", yrs(lambda c: f'=IF({c}{R["int_total_cash"]}>0.05,{c}{R["cm_ebitda"]}/{c}{R["int_total_cash"]},"n.m.")', PROJ), fmt=MULT)
line("cm_icr_capex", "(EBITDA - capex) / cash interest", yrs(lambda c: f'=IF({c}{R["int_total_cash"]}>0.05,({c}{R["cm_ebitda"]}+{c}{R["cf_capex"]})/{c}{R["int_total_cash"]},"n.m.")', PROJ), fmt=MULT)
line("cm_dscr", "EBITDA / (cash interest + mandatory amortisation)", yrs(lambda c: f'=IF({c}{R["int_total_cash"]}-{c}{R["ds_mand"]}>0.05,{c}{R["cm_ebitda"]}/({c}{R["int_total_cash"]}-{c}{R["ds_mand"]}),"n.m.")', PROJ), fmt=MULT)
line("cm_head_lev", "Headroom to max net leverage covenant (x)", yrs(lambda c: f"=$D${R['cov_lev']}-{c}{R['cm_nlev']}", PROJ), fmt=MULT2)
line("cm_head_icr", "Headroom to min interest cover covenant (x)", yrs(lambda c: f'=IF(ISNUMBER({c}{R["cm_icr"]}),{c}{R["cm_icr"]}-$D${R["cov_icr"]},"n.m.")', PROJ), fmt=MULT)
line("cm_max_nlev", "Peak net debt / EBITDA over hold", d=f"=MAX(I{R['cm_nlev']}:M{R['cm_nlev']})", fmt=MULT2, bold=True)
line("cm_min_icr", "Minimum EBITDA / cash interest over hold", d=f'=IF(COUNT(I{R["cm_icr"]}:M{R["cm_icr"]})=0,"n.m.",MIN(I{R["cm_icr"]}:M{R["cm_icr"]}))', fmt=MULT, bold=True)
line("cm_repaid", "Year in which funded debt is fully repaid", d=f'=IF(H{R["bs_debt"]}<0.01,"no debt",IF(M{R["bs_debt"]}>0.01,"beyond Y5",COUNTIF(I{R["bs_debt"]}:M{R["bs_debt"]},">0.01")+1))', fmt=INTF)

# ---- returns ----
section("Returns analysis")
yh = nr()
for c, y in zip(["H"] + PROJ, ["Close"] + [f"Y{i} {y}" for i, y in zip(range(1, 6), YEARS[2:])]):
    put(ws, f"{c}{yh}", y, bold=True, align="right")
adv(1)
line("rt_years", "Years since close", yrs(lambda c: f"={c}{R['years']}", PROJ), fmt=INTF)
line("rt_ebitda", "LTM Adj. EBITDA", yrs(lambda c: f"={c}{R['adj_ebitda']}", PROJ))
line("rt_mult", "Exit multiple", yrs(lambda c: f"=$D${R['exit_mult']}", PROJ), fmt=MULT)
line("rt_tev", "Exit TEV", yrs(lambda c: f"={c}{R['rt_ebitda']}*{c}{R['rt_mult']}", PROJ))
line("rt_debt", "Less: debt", yrs(lambda c: f"=-{c}{R['bs_debt']}", PROJ))
line("rt_cash", "Plus: cash", yrs(lambda c: f"={c}{R['bs_cash']}", PROJ))
line("rt_costs", "Less: exit costs", yrs(lambda c: f"=-{c}{R['rt_tev']}*$D${R['exit_cost_pct']}", PROJ))
line("rt_eq", "Equity value at exit", yrs(lambda c: f"=SUM({c}{R['rt_tev']}:{c}{R['rt_costs']})", PROJ), bold=True, top=True)
line("rt_dist", "Distributions received in year", yrs(lambda c: f"=-{c}{R['cf_dist']}", PROJ))
line("rt_cumdist", "Cumulative distributions", yrs(lambda c: f"=SUM($I{R['rt_dist']}:{c}{R['rt_dist']})", PROJ))
line("rt_inv", "Equity invested at close", {"H": f"=D{R['eq_total']}"})
line("rt_moic", "MoIC (exit in year)", yrs(lambda c: f"=({c}{R['rt_eq']}+{c}{R['rt_cumdist']})/$H${R['rt_inv']}", PROJ), fmt=MULT2, bold=True)
irr_rows = {}
for n in range(1, 6):
    rr = nr()
    irr_rows[n] = rr
    put(ws, f"B{rr}", f"  IRR cash flows - exit in Y{n}", italic=True, color=GRAY, size=8)
    put(ws, f"H{rr}", f"=-$H${R['rt_inv']}", fmt=M1, italic=True, size=8)
    for t, c in enumerate(PROJ, start=1):
        if t < n:
            f = f"={c}{R['rt_dist']}"
        elif t == n:
            f = f"={c}{R['rt_dist']}+{c}{R['rt_eq']}"
        else:
            f = None
        if f:
            put(ws, f"{c}{rr}", f, fmt=M1, italic=True, size=8)
    adv(1)
line("rt_irr", "IRR (exit in year)", {c: "=" + irr_f(f"$H{irr_rows[i]}:{c}{irr_rows[i]}") for i, c in enumerate(PROJ, start=1)}, fmt=PCT, bold=True)
line("out_moic", "HEADLINE: 5-year MoIC", d=f"=M{R['rt_moic']}", fmt=MULT2, bold=True, fill=F_KEY)
ws[f"D{R['out_moic']}"].border = BOX
line("out_irr", "HEADLINE: 5-year IRR", d=f"=M{R['rt_irr']}", fmt=PCT, bold=True, fill=F_KEY)
ws[f"D{R['out_irr']}"].border = BOX
line("out_eq", "Equity value at exit (Y5) + cumulative distributions", d=f"=M{R['rt_eq']}+M{R['rt_cumdist']}")
name("OUT_IRR", f"LBO!$D${R['out_irr']}")
name("OUT_MOIC", f"LBO!$D${R['out_moic']}")

# exit multiple x exit year grid
adv(1)
gh = nr()
put(ws, f"B{gh}", "Exit multiple sensitivity (exit in Y3 / Y4 / Y5)", bold=True)
for c, y in zip(["K", "L", "M"], ["Y3 FY30E", "Y4 FY31E", "Y5 FY32E"]):
    put(ws, f"{c}{gh}", y, bold=True, align="right")
put(ws, f"C{gh}", "MoIC (upper) and IRR (lower)", italic=True, color=GRAY, size=8)
adv(1)
GRID_M = [3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]
R["grid_moic0"] = nr()
for m in GRID_M:
    rr = nr()
    put(ws, f"J{rr}", m, fmt=MULT)
    for c in ["K", "L", "M"]:
        put(ws, f"{c}{rr}", f"=({c}{R['rt_ebitda']}*$J{rr}*(1-$D${R['exit_cost_pct']})-{c}{R['bs_debt']}+{c}{R['bs_cash']}+{c}{R['rt_cumdist']})/$H${R['rt_inv']}", fmt=MULT2)
    put(ws, f"B{rr}", "  MoIC at exit multiple" if m == GRID_M[0] else None)
    adv(1)
adv(1)
R["grid_irr0"] = nr()
helper_start = R["grid_irr0"] + len(GRID_M) + 2
hrow = helper_start
grid_irr_rows = []
for gi, m in enumerate(GRID_M):
    rr = nr()
    grid_irr_rows.append(rr)
    put(ws, f"J{rr}", m, fmt=MULT)
    if gi == 0:
        put(ws, f"B{rr}", "  IRR at exit multiple")
    adv(1)
adv(1)
put(ws, f"B{nr()}", "  IRR helper cash flows by exit multiple and exit year", italic=True, color=GRAY, size=8)
adv(1)
for gi, m in enumerate(GRID_M):
    for n, col in zip([3, 4, 5], ["K", "L", "M"]):
        rr = nr()
        put(ws, f"B{rr}", f"    helper {m:.1f}x / exit Y{n}", italic=True, color=GRAY, size=7)
        put(ws, f"G{rr}", f"=J{grid_irr_rows[gi]}", fmt=MULT, italic=True, size=7)
        put(ws, f"H{rr}", f"=-$H${R['rt_inv']}", fmt=M1, italic=True, size=7)
        for t, c in enumerate(PROJ, start=1):
            if t < n:
                f = f"={c}{R['rt_dist']}"
            elif t == n:
                f = (f"={c}{R['rt_dist']}+{c}{R['rt_ebitda']}*$G{rr}*(1-$D${R['exit_cost_pct']})"
                     f"-{c}{R['bs_debt']}+{c}{R['bs_cash']}")
            else:
                f = None
            if f:
                put(ws, f"{c}{rr}", f, fmt=M1, italic=True, size=7)
        put(ws, f"{col}{grid_irr_rows[gi]}", "=" + irr_f(f"$H{rr}:{PROJ[n - 1]}{rr}"), fmt=PCT)
        adv(1)

# resolve deferred {key} placeholders in formulas
import re


def resolve(ws, R):
    pat = re.compile(r"\{([a-z_0-9]+)\}")
    for rowc in ws.iter_rows():
        for c in rowc:
            if isinstance(c.value, str) and c.value.startswith("=") and "{" in c.value:
                def rep(mo, cell=c):
                    k = mo.group(1)
                    if k == "r":
                        return str(cell.row)
                    return str(R[k])
                c.value = pat.sub(rep, c.value)


resolve(ws, R)
for rr in range(1, nr() + 2):
    ws.row_dimensions[rr].height = 12.75
ws.freeze_panes = "D5"

# conditional formats on checks (red if non-zero)
for key in ["su_check", "pf_check", "bs_check", "ni_check"]:
    rr = R[key]
    rng = f"D{rr}:M{rr}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(D{rr}),ABS(D{rr})>0.001)'],
                                                   fill=PatternFill("solid", start_color="FFC7CE"),
                                                   font=Font(color="9C0006")))

# ========================================================================================
# ATP
# ========================================================================================
ws = S["ATP"]
title(ws, "Ability to pay (reverse LBO) - running case",
      "Max offer per share that still earns the hurdle IRR, given the running operating/financing case (Y5 exit). "
      "Cash flows include interim distributions.", "K")
widths(ws, {"A": 2, "B": 52, "C": 12, "D": 12, "E": 12, "F": 2, "G": 12, "H": 12, "I": 12, "J": 2, "K": 50})
L = "LBO!"
atp = {}
r = 5
for k, lab, f, fmt in [
    ("ebitda5", "FY32E LTM Adj. EBITDA", f"={L}M{R['adj_ebitda']}", M1),
    ("debt5", "FY32E debt", f"={L}M{R['bs_debt']}", M1),
    ("cash5", "FY32E cash", f"={L}M{R['bs_cash']}", M1),
    ("xc", "Exit costs (% TEV)", f"={L}D{R['exit_cost_pct']}", PCT),
    ("npv_note", "Interim distributions Y1-Y5 are discounted at each hurdle", None, None),
    ("cash_applied", "Clinuvel cash applied at close", f"={L}D{R['cash_applied']}", M1),
    ("debt0", "New debt raised at close", f"={L}D{R['s_rcf']}+{L}D{R['s_tl']}+{L}D{R['s_pik']}", M1),
    ("ff", "Financing fees", f"={L}D{R['u_fin']}", M1),
    ("te", "Transaction expenses (% of equity price)", f"={L}D{R['trans_pct']}", PCT2),
    ("fd", "Fully diluted shares (m)", f"={L}D{R['shares_fd']}", SH),
    ("und", "Undisturbed price (A$)", f"={L}D{R['undisturbed']}", PS),
    ("cashpre", "Pre-deal cash at close", f"={L}D{R['cash_pre']}", M1),
    ("ltm", "LTM (FY27E) Adj. EBITDA", f"={L}D{R['ebitda_ltm']}", M1),
]:
    atp[k] = r
    put(ws, f"B{r}", lab, italic=(f is None), color=GRAY if f is None else BLACK)
    if f:
        put(ws, f"C{r}", f, fmt=fmt)
    r += 1
r += 1
HURDLES = [0.15, 0.20, 0.25]
blocks = [("Max equity invested (A$m)", "inv", M1), ("Max offer price per share (A$)", "px", PS),
          ("Implied premium to undisturbed", "prem", PCT), ("Implied TEV / LTM Adj. EBITDA", "mult", MULT)]
ATP_ROWS = {}
dist_rng = f"{L}I{R['rt_dist']}:L{R['rt_dist']}"
for title_txt, key, fmt in blocks:
    bar(ws, r, title_txt, "B", "I", fill=F_OUT, color=WHITE)
    for c, h in zip("CDE", HURDLES):
        put(ws, f"{c}{r}", h, fmt='0%" IRR"', bold=True, color=WHITE, fill=F_OUT, align="right")
    r += 1
    ATP_ROWS[key] = r
    for m in GRID_M:
        put(ws, f"B{r}", f"Exit at {m:.1f}x LTM EBITDA")
        put(ws, f"G{r}", m, fmt=MULT)
        for c in "CDE":
            hr = ATP_ROWS[key] - 1
            eq_exit = f"($G{r}*$C${atp['ebitda5']}*(1-$C${atp['xc']})-$C${atp['debt5']}+$C${atp['cash5']})"
            inv = f"(NPV({c}${hr},{dist_rng})+({L}M{R['rt_dist']}+{eq_exit})/(1+{c}${hr})^5)"
            if key == "inv":
                f = f"={inv}"
            elif key == "px":
                f = f"=({c}{ATP_ROWS['inv'] + (r - ATP_ROWS[key])}+$C${atp['cash_applied']}+$C${atp['debt0']}-$C${atp['ff']})/(1+$C${atp['te']})/$C${atp['fd']}"
            elif key == "prem":
                f = f"={c}{ATP_ROWS['px'] + (r - ATP_ROWS[key])}/$C${atp['und']}-1"
            else:
                f = f"=({c}{ATP_ROWS['px'] + (r - ATP_ROWS[key])}*$C${atp['fd']}-$C${atp['cashpre']})/$C${atp['ltm']}"
            put(ws, f"{c}{r}", f, fmt=fmt)
        r += 1
    r += 1
# running-exit-multiple ATP + circular check
bar(ws, r, "ATP at the running exit multiple & check", "B", "I")
r += 1
atp["run_mult"] = r
put(ws, f"B{r}", "Running exit multiple"); put(ws, f"C{r}", f"={L}D{R['exit_mult']}", fmt=MULT); r += 1
atp["run_rows"] = {}
for h in HURDLES:
    put(ws, f"B{r}", f"Max offer price at {h:.0%} IRR (A$/sh)")
    eq_exit = f"($C${atp['run_mult']}*$C${atp['ebitda5']}*(1-$C${atp['xc']})-$C${atp['debt5']}+$C${atp['cash5']})"
    inv = f"(NPV({h},{dist_rng})+({L}M{R['rt_dist']}+{eq_exit})/(1+{h})^5)"
    put(ws, f"C{r}", f"=({inv}+$C${atp['cash_applied']}+$C${atp['debt0']}-$C${atp['ff']})/(1+$C${atp['te']})/$C${atp['fd']}", fmt=PS, bold=True)
    atp["run_rows"][h] = r
    r += 1
atp["check"] = r
put(ws, f"B{r}", "Check: price implied at the model's own IRR less running offer (should be ~0)", italic=True)
eq_exit = f"($C${atp['run_mult']}*$C${atp['ebitda5']}*(1-$C${atp['xc']})-$C${atp['debt5']}+$C${atp['cash5']})"
irr_ref = f"{L}D{R['out_irr']}"
inv = f"(NPV({irr_ref},{dist_rng})+({L}M{R['rt_dist']}+{eq_exit})/(1+{irr_ref})^5)"
put(ws, f"C{r}", f"=({inv}+$C${atp['cash_applied']}+$C${atp['debt0']}-$C${atp['ff']})/(1+$C${atp['te']})/$C${atp['fd']}-{L}D{R['offer']}",
    fmt='0.0000', italic=True, fill=F_CHECK)
put(ws, "K5", "Max equity = NPV at the hurdle of distributions (Y1-Y4) plus Y5 distribution and exit equity. "
    "Max price = (max equity + cash applied + new debt - financing fees) / (1 + fee%) / FD shares. "
    "Debt and cash applied do not depend on price.", italic=True, color=GRAY, size=8, wrap=True)
set_rows(ws)

# ========================================================================================
# Contrib (value creation bridge)
# ========================================================================================
ws = S["Contrib"]
title(ws, "Value creation bridge - running case (entry at close to Y5 exit)",
      "Decomposes (exit equity + distributions - equity invested). Entry multiple is on status-quo FY27E EBITDA, "
      "so the EBITDA bucket includes the sponsor cost-out.", "F")
widths(ws, {"A": 2, "B": 58, "C": 12, "D": 12, "E": 2, "F": 60})
cb = {}
r = 5
items = [
    ("tev0", "Entry TEV (implied by offer)", f"={L}D{R['tev']}", M1),
    ("e0", "Entry LTM Adj. EBITDA (FY27E)", f"={L}H{R['adj_ebitda']}", M1),
    ("m0", "Entry multiple", "=C{tev0}/C{e0}", MULT2),
    ("rev0", "Entry revenue (FY27E)", f"={L}H{R['rev_total']}", M1),
    ("mg0", "Entry Adj. EBITDA margin", "=C{e0}/C{rev0}", PCT),
    ("e1", "Exit LTM Adj. EBITDA (FY32E)", f"={L}M{R['adj_ebitda']}", M1),
    ("m1", "Exit multiple", f"={L}D{R['exit_mult']}", MULT2),
    ("tev1", "Exit TEV", f"={L}M{R['rt_tev']}", M1),
    ("rev1", "Exit revenue (FY32E)", f"={L}M{R['rev_total']}", M1),
    ("mg1", "Exit Adj. EBITDA margin", "=C{e1}/C{rev1}", PCT),
    ("nd0", "PF net debt at close", f"={L}I{R['pf_debt']}-{L}I{R['pf_cash']}", M1),
    ("nd1", "Net debt at exit", f"={L}M{R['bs_debt']}-{L}M{R['bs_cash']}", M1),
    ("dist", "Cumulative distributions", f"={L}M{R['rt_cumdist']}", M1),
    ("fees", "Entry fees (transaction + financing)", f"=-({L}D{R['u_trans']}+{L}D{R['u_fin']})", M1),
    ("xcost", "Exit costs", f"={L}M{R['rt_costs']}", M1),
    ("inv", "Equity invested", f"={L}D{R['eq_total']}", M1),
    ("eq1", "Exit equity value", f"={L}M{R['rt_eq']}", M1),
]
for k, lab, f, fmt in items:
    cb[k] = r
    r += 1
for k, lab, f, fmt in items:
    rr = cb[k]
    put(ws, f"B{rr}", lab)
    put(ws, f"C{rr}", re.sub(r"\{([a-z0-9]+)\}", lambda mo: str(cb[mo.group(1)]), f), fmt=fmt)
r += 1
bar(ws, r, "Bridge", "B", "D", fill=F_OUT, color=WHITE)
put(ws, f"C{r}", "A$m", bold=True, color=WHITE, fill=F_OUT, align="right")
put(ws, f"D{r}", "% of total", bold=True, color=WHITE, fill=F_OUT, align="right")
r += 1
br = {}
bridge = [
    ("b_rev", "EBITDA growth: revenue effect  [(Rev1-Rev0) x margin0 x m0]", f"=(C{cb['rev1']}-C{cb['rev0']})*C{cb['mg0']}*C{cb['m0']}"),
    ("b_mgn", "EBITDA growth: margin effect  [(mgn1-mgn0) x Rev1 x m0]", f"=(C{cb['mg1']}-C{cb['mg0']})*C{cb['rev1']}*C{cb['m0']}"),
    ("b_mult", "Multiple expansion / (contraction)  [(m1-m0) x EBITDA1]", f"=(C{cb['m1']}-C{cb['m0']})*C{cb['e1']}"),
    ("b_cash", "Deleveraging & cash generation  [ND0 - ND1 + distributions]", f"=C{cb['nd0']}-C{cb['nd1']}+C{cb['dist']}"),
    ("b_fees", "Entry fees", f"=C{cb['fees']}"),
    ("b_xc", "Exit costs", f"=C{cb['xcost']}"),
]
for k, lab, f in bridge:
    br[k] = r
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f, fmt=M1)
    r += 1
br["total"] = r
put(ws, f"B{r}", "Total value created", bold=True)
put(ws, f"C{r}", f"=SUM(C{br['b_rev']}:C{br['b_xc']})", fmt=M1, bold=True, border=TOPLINE)
r += 1
for k, _, _ in bridge:
    put(ws, f"D{br[k]}", f"=C{br[k]}/$C${br['total']}", fmt=PCT)
put(ws, f"D{br['total']}", f"=C{br['total']}/$C${br['total']}", fmt=PCT, bold=True)
br["check"] = r
put(ws, f"B{r}", "Check: exit equity + distributions - invested - bridge (should be 0)", italic=True)
put(ws, f"C{r}", f"=C{cb['eq1']}+C{cb['dist']}-C{cb['inv']}-C{br['total']}", fmt='0.000', italic=True, fill=F_CHECK)
put(ws, "F5", "Reading the bridge: a declining single-product asset should show negative revenue effect and "
    "multiple contraction, offset by cost-out (margin effect) and cash generation. If most value comes from "
    "cash generation, the deal is a 'cash-box' harvest rather than an operational growth story.",
    italic=True, color=GRAY, size=8, wrap=True)
set_rows(ws)

# ========================================================================================
# Sensitivity
# ========================================================================================
ws = S["Sensitivity"]
title(ws, "Sensitivity & case outputs",
      "Block A-B: live formulas (running case; helper cash flows in columns AL-AQ). Block C: self-referencing capture cells - run every MCASE "
      "(model/run_cases.py or cycle MCASE in Excel with iteration on). Blocks D-F read the captured values.", "AB")
widths(ws, {"A": 2, "B": 8, "C": 34})
for ci in range(4, 29):
    ws.column_dimensions[get_column_letter(ci)].width = 10.5
r = 5
# ---- Block A/B: offer price x exit multiple (running case, Y5) ----
PRICES = [9.00, 9.50, 10.00, 10.50, 11.00, 11.50, 12.00, 12.50, 13.00]
sens_mults = [3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]
inv_expr = lambda pcell: (f"({pcell}*{L}$D${R['shares_fd']}*(1+{L}$D${R['trans_pct']})+{L}$D${R['non_price_inv']})")
bar(ws, r, "Block A: IRR - offer price (rows) x exit multiple (columns), running operating & financing case, exit Y5", "B", "L")
r += 1
put(ws, f"C{r}", "Offer (A$/sh) | premium", bold=True)
for j, m in enumerate(sens_mults):
    put(ws, f"{get_column_letter(5 + j)}{r}", m, fmt=MULT, bold=True, align="right")
a_hdr = r
r += 1
a_rows = []
for p in PRICES:
    put(ws, f"C{r}", p, fmt=PS)
    put(ws, f"D{r}", f"=C{r}/{L}$D${R['undisturbed']}-1", fmt=PCT, italic=True)
    a_rows.append(r)
    r += 1
r += 1
bar(ws, r, "Block B: MoIC - offer price x exit multiple (running case, exit Y5)", "B", "L")
r += 1
b_hdr = r
for j, m in enumerate(sens_mults):
    put(ws, f"{get_column_letter(5 + j)}{r}", m, fmt=MULT, bold=True, align="right")
r += 1
b_rows = []
for p in PRICES:
    put(ws, f"C{r}", p, fmt=PS)
    put(ws, f"D{r}", f"=C{r}/{L}$D${R['undisturbed']}-1", fmt=PCT, italic=True)
    b_rows.append(r)
    r += 1
r += 1
# helper rows for Block A IRRs (price x multiple) placed far right (columns T..Z)
put(ws, "AL4", "Block A helper: cash flows (t=0..5) per price x multiple", italic=True, color=GRAY, size=8)
hr = 5
dist_cells = [f"{L}{c}${R['rt_dist']}" for c in PROJ]
exit_eq = lambda mcell: (f"({L}$M${R['rt_ebitda']}*{mcell}*(1-{L}$D${R['exit_cost_pct']})-{L}$M${R['bs_debt']}+{L}$M${R['bs_cash']})")
for i, prow in enumerate(a_rows):
    for j, m in enumerate(sens_mults):
        mcell = f"{get_column_letter(5 + j)}${a_hdr}"
        put(ws, f"AL{hr}", f"=-{inv_expr(f'$C${prow}')}", fmt=M1, size=7, italic=True)
        for t in range(1, 6):
            col = get_column_letter(38 + t)
            f = f"={dist_cells[t - 1]}" if t < 5 else f"={dist_cells[4]}+{exit_eq(mcell)}"
            put(ws, f"{col}{hr}", f, fmt=M1, size=7, italic=True)
        put(ws, f"{get_column_letter(5 + j)}{prow}", '=IF(-AL{0}<=0,"n.m.",{1})'.format(hr, irr_f(f"AL{hr}:AQ{hr}")), fmt=PCT)
        put(ws, f"{get_column_letter(5 + j)}{b_rows[i]}",
            f'=IF(-AL{hr}<=0,"n.m.",(SUM(AM{hr}:AQ{hr}))/(-AL{hr}))', fmt=MULT2)
        hr += 1

# ---- Block C: case capture (self-referencing IF) ----
bar(ws, r, "Block C: case capture table (self-referencing IF: each row snapshots the model when MCASE = its case #)", "B", "AB")
r += 1
CAP = [
    ("offer", "Offer A$/sh", f"{L}$D${R['offer']}", PS),
    ("prem", "Premium", f"{L}$D${R['premium']}", PCT),
    ("tev_ltm", "TEV / LTM EBITDA", f"{L}$D${R['tev_ltm_x']}", MULT),
    ("tev_fy28", "TEV / FY28E EBITDA", f"{L}$D${R['tev_fy28_x']}", MULT),
    ("eq", "Equity invested", f"{L}$D${R['eq_total']}", M1),
    ("eqpct", "Equity % (eq+new debt)", f"{L}$D${R['eq_pct_new']}", PCT),
    ("debt0", "New debt", f"{L}$D${R['s_tl']}+{L}$D${R['s_pik']}", M1),
    ("rev28", "FY28E revenue", f"{L}$I${R['rev_total']}", M1),
    ("rev29s", "FY29E SCENESSE revenue", f"{L}$J${R['rev_us']}+{L}$J${R['rev_eu']}", M1),
    ("rev32", "FY32E revenue", f"{L}$M${R['rev_total']}", M1),
    ("e27", "FY27E Adj. EBITDA", f"{L}$H${R['adj_ebitda']}", M1),
    ("e28", "FY28E Adj. EBITDA", f"{L}$I${R['adj_ebitda']}", M1),
    ("e32", "FY32E Adj. EBITDA", f"{L}$M${R['adj_ebitda']}", M1),
    ("cumdist", "Cum. distributions", f"{L}$M${R['rt_cumdist']}", M1),
    ("eqx", "Exit equity (Y5)", f"{L}$M${R['rt_eq']}", M1),
    ("moic", "MoIC (Y5)", f"{L}$D${R['out_moic']}", MULT2),
    ("irr", "IRR (Y5)", f"{L}$D${R['out_irr']}", PCT),
    ("maxlev", "Peak net lev.", f"{L}$D${R['cm_max_nlev']}", MULT2),
    ("minicr", "Min EBITDA/interest", f"{L}$D${R['cm_min_icr']}", MULT),
    ("atp20", "ATP @20% (A$/sh)", f"ATP!$C${atp['run_rows'][0.20]}", PS),
    ("atp25", "ATP @25% (A$/sh)", f"ATP!$C${atp['run_rows'][0.25]}", PS),
    ("npi", "Non-price inv. component", f"{L}$D${R['non_price_inv']}", M1),
    ("d1", "Dist Y1", f"{L}$I${R['rt_dist']}", M1),
    ("d2", "Dist Y2", f"{L}$J${R['rt_dist']}", M1),
    ("d3", "Dist Y3", f"{L}$K${R['rt_dist']}", M1),
    ("d4", "Dist Y4", f"{L}$L${R['rt_dist']}", M1),
    ("d5", "Dist Y5", f"{L}$M${R['rt_dist']}", M1),
]
put(ws, f"B{r}", "Case", bold=True, color=WHITE, fill=F_OUT)
put(ws, f"C{r}", "Label", bold=True, color=WHITE, fill=F_OUT)
CAPCOL = {}
for j, (k, lab, ref, fmt) in enumerate(CAP):
    col = get_column_letter(4 + j)
    CAPCOL[k] = col
    put(ws, f"{col}{r}", lab, bold=True, color=WHITE, fill=F_OUT, wrap=True, align="center")
ws.row_dimensions[r].height = 36
cap_hdr = r
r += 1
CAPROW = {}
for n, lab, *_ in MCASES:
    CAPROW[n] = r
    put(ws, f"B{r}", n, fmt=INTF, color=BLACK)
    put(ws, f"C{r}", f"=MCASE!C{8 + n}", size=8)
    for k, labx, ref, fmt in CAP:
        col = CAPCOL[k]
        put(ws, f"{col}{r}", f"=IF(MCASE=$B{r},{ref},{col}{r})", fmt=fmt)
    r += 1
r += 1
# ---- Block D: op case x fin case matrices ----
bar(ws, r, "Block D: IRR and MoIC by operating case (rows) x financing case (columns) at the base offer", "B", "L")
r += 1
for metric, fmt in [("irr", PCT), ("moic", MULT2), ("eq", M1)]:
    put(ws, f"C{r}", {"irr": "IRR (Y5)", "moic": "MoIC (Y5)", "eq": "Equity invested (A$m)"}[metric], bold=True)
    for j, (fn, fnm, _) in enumerate(FINCASES):
        put(ws, f"{get_column_letter(4 + j)}{r}", fnm.split(" (")[0], bold=True, align="right", wrap=True)
    ws.row_dimensions[r].height = 24
    r += 1
    for i, (on, onm, _, _) in enumerate(CORE_OPS):
        put(ws, f"C{r}", f"{on}. {onm}")
        for j in range(4):
            cn = j * 5 + on
            put(ws, f"{get_column_letter(4 + j)}{r}", f"={CAPCOL[metric]}{CAPROW[cn]}", fmt=fmt)
        r += 1
    r += 1
# ---- Block E: probability-weighted view (base financing) ----
bar(ws, r, "Block E: probability-weighted outcome (base financing, base offer) - illustrative weights", "B", "L")
r += 1
put(ws, f"C{r}", "Operating case", bold=True)
for col, h in zip("DEFGH", ["Probability", "IRR", "MoIC", "Exit eq.+dist.", "ATP @20%"]):
    put(ws, f"{col}{r}", h, bold=True, align="right")
r += 1
e_first = r
for i, (on, onm, _, _) in enumerate(CORE_OPS):
    put(ws, f"C{r}", f"{on}. {onm}")
    put(ws, f"D{r}", CASE_PROB[i], fmt=PCT)
    put(ws, f"E{r}", f"={CAPCOL['irr']}{CAPROW[on]}", fmt=PCT)
    put(ws, f"F{r}", f"={CAPCOL['moic']}{CAPROW[on]}", fmt=MULT2)
    put(ws, f"G{r}", f"={CAPCOL['eqx']}{CAPROW[on]}+{CAPCOL['cumdist']}{CAPROW[on]}", fmt=M1)
    put(ws, f"H{r}", f"={CAPCOL['atp20']}{CAPROW[on]}", fmt=PS)
    r += 1
e_last = r - 1
put(ws, f"C{r}", "Probability-weighted", bold=True)
put(ws, f"D{r}", f"=SUM(D{e_first}:D{e_last})", fmt=PCT, bold=True)
put(ws, f"E{r}", f"=SUMPRODUCT($D{e_first}:$D{e_last},E{e_first}:E{e_last})", fmt=PCT, bold=True)
put(ws, f"F{r}", f"=SUMPRODUCT($D{e_first}:$D{e_last},F{e_first}:F{e_last})", fmt=MULT2, bold=True)
put(ws, f"G{r}", f"=SUMPRODUCT($D{e_first}:$D{e_last},G{e_first}:G{e_last})", fmt=M1, bold=True)
put(ws, f"H{r}", f"=SUMPRODUCT($D{e_first}:$D{e_last},H{e_first}:H{e_last})", fmt=PS, bold=True)
e_tot = r
put(ws, f"I{r}", "(IRR/MoIC here = probability-weighted averages)", italic=True, color=GRAY, size=8)
r += 1
put(ws, f"C{r}", "Expected cash flows to equity, t = 0..5 (A$m)", bold=True)
cap_rng = lambda k: f"{CAPCOL[k]}{CAPROW[1]}:{CAPCOL[k]}{CAPROW[5]}"
pr = f"$D${e_first}:$D${e_last}"
put(ws, f"D{r}", f"=-SUMPRODUCT({pr},{cap_rng('eq')})", fmt=M1)
for t in range(1, 6):
    col = get_column_letter(4 + t)
    f = f"=SUMPRODUCT({pr},{cap_rng(f'd{t}')})" + (f"+SUMPRODUCT({pr},{cap_rng('eqx')})" if t == 5 else "")
    put(ws, f"{col}{r}", f, fmt=M1)
e_cf = r
r += 1
put(ws, f"C{r}", "IRR / MoIC of expected cash flows", bold=True)
put(ws, f"D{r}", "=" + irr_f(f"D{e_cf}:I{e_cf}"), fmt=PCT, bold=True, fill=F_KEY)
put(ws, f"E{r}", f"=SUM(E{e_cf}:I{e_cf})/-D{e_cf}", fmt=MULT2, bold=True, fill=F_KEY)
e_irr = r
r += 2
# ---- Block F: offer price x operating case IRR (base financing), using captured cash flows ----
bar(ws, r, "Block F: IRR by offer price (rows) x operating case (columns), base financing - from captured case cash flows", "B", "L")
r += 1
f_hdr = r
put(ws, f"C{r}", "Offer (A$/sh) | premium", bold=True)
for j, (on, onm, _, _) in enumerate(CORE_OPS):
    put(ws, f"{get_column_letter(5 + j)}{r}", f"{on}. {onm}", bold=True, align="right", wrap=True)
put(ws, f"{get_column_letter(10)}{r}", "Expected (prob.-wtd flows)", bold=True, align="right", wrap=True)
ws.row_dimensions[r].height = 24
r += 1
f_rows = []
for p in PRICES:
    put(ws, f"C{r}", p, fmt=PS)
    put(ws, f"D{r}", f"=C{r}/{L}$D${R['undisturbed']}-1", fmt=PCT, italic=True)
    f_rows.append(r)
    r += 1
# helper cash flows for Block F in columns AD..AI
put(ws, "AS4", "Block F helper: cash flows per price x op case (base financing)", italic=True, color=GRAY, size=8)
hr = 5
for i, prow in enumerate(f_rows):
    for j, (on, *_rest) in enumerate(CORE_OPS):
        cr = CAPROW[on]
        inv = f"($C${prow}*{L}$D${R['shares_fd']}*(1+{L}$D${R['trans_pct']})+${CAPCOL['npi']}${cr})"
        put(ws, f"AS{hr}", f"=-{inv}", fmt=M1, size=7, italic=True)
        for t in range(1, 6):
            col = get_column_letter(45 + t)
            dcol = CAPCOL[f"d{t}"]
            f = f"=${dcol}${cr}" if t < 5 else f"=${dcol}${cr}+${CAPCOL['eqx']}${cr}"
            put(ws, f"{col}{hr}", f, fmt=M1, size=7, italic=True)
        put(ws, f"{get_column_letter(5 + j)}{prow}", '=IF(-AS{0}<=0,"n.m.",{1})'.format(hr, irr_f(f"AS{hr}:AX{hr}")), fmt=PCT)
        hr += 1
    # expected cash flows across the five cases at this price -> IRR (helper columns AZ..BE)
    t0 = "+".join(f"$D${e_first + j}*($C${prow}*{L}$D${R['shares_fd']}*(1+{L}$D${R['trans_pct']})+${CAPCOL['npi']}${CAPROW[j + 1]})" for j in range(5))
    put(ws, f"AZ{prow}", f"=-({t0})", fmt=M1, size=7, italic=True)
    for t in range(1, 6):
        col = get_column_letter(52 + t)
        dcol = CAPCOL[f"d{t}"]
        f = f"=SUMPRODUCT($D${e_first}:$D${e_last},${dcol}${CAPROW[1]}:${dcol}${CAPROW[5]})"
        if t == 5:
            f += f"+SUMPRODUCT($D${e_first}:$D${e_last},${CAPCOL['eqx']}${CAPROW[1]}:${CAPCOL['eqx']}${CAPROW[5]})"
        put(ws, f"{col}{prow}", f, fmt=M1, size=7, italic=True)
    put(ws, f"J{prow}", "=" + irr_f(f"AZ{prow}:BE{prow}"), fmt=PCT, bold=True)
r += 1
put(ws, f"C{r}", "Note: IRRs include interim distributions (cash above minimum cash once debt is repaid). "
    "Case exit multiples: " + ", ".join(f"{OPCASES[i][1]} {EXIT_MULT[i + 1]:.1f}x" for i in range(5)) + ".",
    italic=True, color=GRAY, size=8)
set_rows(ws, 12.75)
ws.row_dimensions[cap_hdr].height = 36
ws.row_dimensions[f_hdr].height = 24
ws.freeze_panes = "D5"

# ========================================================================================
# CVR
# ========================================================================================
ws = S["CVR"]
title(ws, "Contingent value rights (CVR) - bridging the bid-ask on competition and vitiligo",
      "Illustrative. CVR 1 pays if SCENESSE holds up (FY29 EPP revenue threshold); CVR 2 pays on vitiligo approval.", "H")
widths(ws, {"A": 2, "B": 60, "C": 12, "D": 12, "E": 12, "F": 2, "G": 2, "H": 60})
cv = {}
r = 5
bar(ws, r, "Inputs", "B", "E")
r += 1
cvr_inputs = [
    ("fd", "Fully diluted shares (m)", f"={L}D{R['shares_fd']}", SH),
    ("offer", "Cash offer per share (A$)", f"={L}D{R['offer']}", PS),
    ("c1_amt", "CVR 1: payment per share (A$) if FY29E SCENESSE revenue >= threshold", 1.00, PS),
    ("c1_thr", "CVR 1: FY29E SCENESSE revenue threshold (A$m)", 90.0, M1),
    ("c1_t", "CVR 1: years from close to payment", 2.25, '0.00'),
    ("c2_amt", "CVR 2: payment per share (A$) on FDA/EC vitiligo approval by 30-Jun-2030", 1.00, PS),
    ("c2_p", "CVR 2: probability of approval by 30-Jun-2030 (unconditional, incl. CUV105 risk)", 0.30, PCT),
    ("c2_t", "CVR 2: years from close to payment", 3.0, '0.00'),
    ("disc", "Discount rate for CVR valuation", 0.12, PCT),
]
for k, lab, v, fmt in cvr_inputs:
    cv[k] = r
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", v, fmt=fmt)
    r += 1
r += 1
bar(ws, r, "CVR 1 trigger by operating case (captured FY29E SCENESSE revenue)", "B", "E")
r += 1
put(ws, f"B{r}", "Operating case", bold=True)
for col, h in zip("CDE", ["Prob.", "FY29E SCEN.", "Triggers?"]):
    put(ws, f"{col}{r}", h, bold=True, align="right")
r += 1
c1_first = r
for i, (on, onm, _, _) in enumerate(CORE_OPS):
    put(ws, f"B{r}", f"{on}. {onm}")
    put(ws, f"C{r}", f"=Sensitivity!D{e_first + i}", fmt=PCT)
    put(ws, f"D{r}", f"=Sensitivity!{CAPCOL['rev29s']}{CAPROW[on]}", fmt=M1)
    put(ws, f"E{r}", f"=IF(D{r}>=$C${cv['c1_thr']},1,0)", fmt=INTF)
    r += 1
c1_last = r - 1
r += 1
bar(ws, r, "Valuation", "B", "E")
r += 1
out = [
    ("c1_p", "CVR 1 probability (sum of triggering case weights)", f"=SUMPRODUCT(C{c1_first}:C{c1_last},E{c1_first}:E{c1_last})", PCT),
    ("c1_pv", "CVR 1 expected PV per share (A$)", "=C{c1_amt}*C{c1_p}/(1+C{disc})^C{c1_t}", PS),
    ("c2_pv", "CVR 2 expected PV per share (A$)", "=C{c2_amt}*C{c2_p}/(1+C{disc})^C{c2_t}", PS),
    ("face", "Headline value per share: cash + CVR face (A$)", "=C{offer}+C{c1_amt}+C{c2_amt}", PS),
    ("riskadj", "Risk-adjusted value per share: cash + expected PV of CVRs (A$)", "=C{offer}+C{c1_pv}+C{c2_pv}", PS),
    ("cost", "Expected PV cost to sponsor (A$m)", "=(C{c1_pv}+C{c2_pv})*C{fd}", M1),
    ("maxcost", "Maximum CVR payout (A$m, undiscounted)", "=(C{c1_amt}+C{c2_amt})*C{fd}", M1),
]
for k, lab, f, fmt in out:
    cv[k] = r
    r += 1
for k, lab, f, fmt in out:
    rr = cv[k]
    put(ws, f"B{rr}", lab, bold=k in ("face", "riskadj"))
    put(ws, f"C{rr}", re.sub(r"\{([a-z0-9_]+)\}", lambda mo: str(cv[mo.group(1)]), f), fmt=fmt, bold=k in ("face", "riskadj"))
put(ws, "H6", "Design logic: CVR 1 transfers the dersimelagon/bitopertin outcome risk back to sellers - it only pays "
    "in the cases where SCENESSE revenue holds up, i.e. when the sponsor can afford it. CVR 2 lets holders keep the "
    "vitiligo option without the sponsor funding CUV107 on spec. CVRs are uncommon in ASX schemes (need IER valuation "
    "and disclosure) but standard in US biotech M&A; non-transferable CVRs avoid US registration.",
    italic=True, color=GRAY, size=8, wrap=True)
set_rows(ws)

# ========================================================================================
# Cover
# ========================================================================================
ws = S["Cover"]
widths(ws, {"A": 2, "B": 48, "C": 16, "D": 16, "E": 60})
title(ws, "Project Daylight - Clinuvel Pharmaceuticals (ASX: CUV / Nasdaq: CUVL)",
      "Private-equity take-private screening model | prepared 22-Sep-2026 from public information | A$ millions | FYE 30 June", "E")
put(ws, "B5", "PROJECT", bold=True); put(ws, "C5", "Daylight", bold=True, color=BLUE)
name("PROJECT", "Cover!$C$5")
put(ws, "B6", "Running case (MCASE)"); put(ws, "C6", "=MCASE", fmt=INTF, color=GREEN)
put(ws, "B7", "Case label"); put(ws, "C7", "=MCASE!C8", color=GREEN)
bar(ws, 9, "Key outputs - running case", "B", "E")
kos = [
    ("Offer price per share (A$)", f"={L}D{R['offer']}", PS),
    ("Premium to undisturbed (A$8.24)", f"={L}D{R['premium']}", PCT),
    ("Equity purchase price (A$m)", f"={L}D{R['eqpp']}", M1),
    ("Implied TEV (A$m)", f"={L}D{R['tev']}", M1),
    ("TEV / FY27E status-quo Adj. EBITDA", f"={L}D{R['tev_ltm_x']}", MULT),
    ("TEV / FY28E post cost-out Adj. EBITDA", f"={L}D{R['tev_fy28_x']}", MULT),
    ("Clinuvel cash applied (A$m)", f"={L}D{R['cash_applied']}", M1),
    ("New debt (A$m)", f"={L}D{R['s_tl']}+{L}D{R['s_pik']}", M1),
    ("Equity invested (A$m)", f"={L}D{R['eq_total']}", M1),
    ("5-year MoIC", f"={L}D{R['out_moic']}", MULT2),
    ("5-year IRR", f"={L}D{R['out_irr']}", PCT),
    ("Max offer at 20% IRR (A$/sh)", f"=ATP!C{atp['run_rows'][0.20]}", PS),
    ("Model checks (0 = OK)", f"=ROUND(ABS({L}D{R['su_check']})+SUMPRODUCT(ABS({L}G{R['bs_check']}:M{R['bs_check']}))+ABS({L}I{R['pf_check']})+ABS(Contrib!C{br['check']})+ABS(ATP!C{atp['check']}),3)", '0.000'),
]
r = 10
for lab, f, fmt in kos:
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f, fmt=fmt, bold=True)
    if lab.startswith("Model checks"):
        name("CHECKS", f"Cover!$C${r}")
    r += 1
r += 1
bar(ws, r, "Tabs", "B", "E")
r += 1
for tab, desc in [
    ("Hist", "Reported history FY20-FY26 and FY26A base-year inputs (reported vs estimated flagged)"),
    ("MCASE", "Master case driver - 23 cases (5 operating x 4 financing at the base offer + 3 variants)"),
    ("OPCASE", "Five operating cases: revenue by region, cost base, sponsor cost-out, capex, NWC"),
    ("FINCASE", "Four financing cases: base 2.0x private credit, high 3.0x + PIK, low 1.0x, no debt"),
    ("LBO", "Single-tab vertical model: S&U, PF BS, IS, BS, CF, debt, interest, credit, returns"),
    ("ATP", "Reverse LBO: max offer per share by exit multiple and hurdle IRR"),
    ("Contrib", "Value creation bridge (revenue, margin, multiple, cash generation, fees)"),
    ("Sensitivity", "Price x multiple grids; case capture table; op x fin matrices; probability-weighted view"),
    ("CVR", "Contingent value right design and valuation"),
]:
    put(ws, f"B{r}", tab, bold=True)
    put(ws, f"C{r}", desc)
    r += 1
r += 1
bar(ws, r, "Colour code & mechanics", "B", "E")
r += 1
for lab, colr in [("Blue: hard-coded input", BLUE), ("Black: formula on same tab", BLACK),
                  ("Green: link to another tab", GREEN)]:
    put(ws, f"B{r}", lab, color=colr)
    r += 1
put(ws, f"B{r}", "Iterative calculation must be ON (File > Options > Formulas; 1,000 iterations, 0.0000001). "
    "Average-balance interest (CIRC=1) and the Sensitivity capture cells are circular by design.", wrap=False, size=8)
r += 1
put(ws, f"B{r}", "To refresh the case tables: set MCASE to 1..23 in turn (or run model/run_cases.py), then back to the base case (3).", size=8)
r += 2
bar(ws, r, "Important", "B", "E")
r += 1
for txt in [
    "Screening model built from public information only (company releases, press, third-party research summaries).",
    "Several inputs are analyst estimates, not disclosures: FY26 cost split by function, US/EU revenue split,",
    "working capital, dilutive securities, financing terms and all forward scenarios. They are flagged in the source notes.",
    "Not investment advice. Figures should be tied to the FY2026 annual report and 20-F before any use.",
]:
    put(ws, f"B{r}", txt, italic=True, size=8)
    r += 1
set_rows(ws)

# ----------------------------------------------------------------------------------------
wb.calculation = CalcProperties(iterate=True, iterateCount=1000, iterateDelta=0.0000001, fullCalcOnLoad=True)
wb.active = 0
wb.save(OUT)
print(f"Saved {OUT}")
print("LBO key rows:", {k: R[k] for k in ["offer", "tev", "eq_total", "adj_ebitda", "out_moic", "out_irr", "bs_check", "pf_check"]})
