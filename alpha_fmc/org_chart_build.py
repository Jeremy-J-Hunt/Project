"""Generate the Alpha FMC org chart page (inline SVG) -> alpha_fmc/org_chart.html"""
from html import escape as esc

W, H = 1200, 1250
out = []


def text_lines(x, y, lines, anchor="middle"):
    """lines: list of (cls, text). Returns svg text elements stacked from y (baseline)."""
    s, cy = [], y
    for cls, t in lines:
        s.append(f'<text x="{x}" y="{cy}" class="{cls}" text-anchor="{anchor}">{esc(t)}</text>')
        cy += 13.5 if cls != "t" else 14.5
    return "".join(s)


def box(x, y, w, h, title, num=None, note=None, cls="", note2=None):
    lines = [("t", t) for t in (title if isinstance(title, list) else [title])]
    if num:
        lines.append(("n", num))
    for n in (note, note2):
        if n:
            lines.append(("m", n))
    block = sum(14.5 if c == "t" else 13.5 for c, _ in lines)
    y0 = y + (h - block) / 2 + 10.5
    out.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" rx="3" class="bx {cls}"/>'
               f'{text_lines(x + w / 2, y0, lines)}</g>')


def tag(x, y, w, h, key, vals):
    out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="2" class="tg"/>')
    lines = [("tk", key)] + [("tv", v) for v in vals]
    block = 13.5 * len(lines)
    out.append(text_lines(x + 10, y + (h - block) / 2 + 10.5, lines, anchor="start"))


def path(pts, cls="e", arrow=True):
    d = "M" + " L".join(f"{a},{b}" for a, b in pts)
    m = ' marker-end="url(#ah)"' if arrow else ""
    out.append(f'<path d="{d}" class="{cls}"{m}/>')


def label(x, y, t, anchor="middle", cls="el"):
    out.append(f'<text x="{x}" y="{y}" class="{cls}" text-anchor="{anchor}">{esc(t)}</text>')


# ---------------- sponsor tier ----------------
box(380, 14, 440, 62, "Bridgepoint Group plc", "11443992",
    "via OP GP → Group Holdings → Advisers Group → Advisers Holdings (01899316)", "sp")
path([(600, 76), (600, 94)], arrow=False)
path([(250, 94), (950, 94)], arrow=False)
for cx in (250, 600, 950):
    path([(cx, 94), (cx, 110)])
box(125, 112, 250, 62, "Bridgepoint Advisers Ltd", "03220373", "fund manager", "sp")
box(475, 112, 250, 62, "Bridgepoint Europe VII GP LLP", "OC438202", "fund GP · via Bridgepoint Europe (SGP) Ltd", "sp")
box(825, 112, 250, 62, "Bridgepoint Europe VII Nominees Ltd", "13793519", "holds shares for the fund", "sp")

# edges into Topco
path([(250, 174), (250, 205), (520, 205), (520, 234)], "e dash")
label(385, 199, "significant influence")
path([(600, 174), (600, 234)])
label(594, 222, "controls >75%", anchor="end")
path([(950, 174), (950, 205), (680, 205), (680, 234)])
label(815, 199, "registered shareholder")

# ---------------- Actium stack (spine) ----------------
SX, SW = 450, 300
rows = {
    "topco": 236, "mid1": 330, "mid2": 410, "mid3": 490,
    "bidco": 570, "hold": 650, "abid": 730, "grp": 810,
}
BH = 56
box(SX, rows["topco"], SW, 62, "Actium Topco (UK) Ltd", "15735478", "group parent · files consolidated accounts", "stk top")
box(SX, rows["mid1"], SW, BH, "Actium Midco 1 (UK) Ltd", "15735608", None, "stk")
box(SX, rows["mid2"], SW, BH, "Actium Midco 2 (UK) Ltd", "15735686", None, "stk")
box(SX, rows["mid3"], SW, BH, "Actium Midco 3 (UK) Ltd", "15736277", None, "stk")
box(SX, rows["bidco"], SW, BH, "Actium Bidco (UK) Ltd", "15736419", "offeror", "stk")
box(SX, rows["hold"], SW, BH, "Actium Holdings Ltd", "09965297", "formerly Alpha Financial Markets Consulting plc", "tgt")
box(SX, rows["abid"], SW, BH, "Alpha FMC Bidco Ltd", "09928343", "legacy 2016 buyout vehicle", "")
box(SX, rows["grp"], SW, BH, "Alpha Financial Markets Consulting Group Ltd", "07160664", "holding company of the trading group", "")

seq = ["topco", "mid1", "mid2", "mid3", "bidco", "hold", "abid", "grp"]
for a, b in zip(seq, seq[1:]):
    top = rows[a] + (62 if a == "topco" else BH)
    path([(600, top), (600, rows[b] - 2)])
    label(607, (top + rows[b]) / 2 + 4, "100%", anchor="start", cls="el sm")

# management into Topco
box(880, 236, 260, 62, "Management + employee trust", None,
    "CSC Employee Benefit Trust (Jersey) Ltd", "mg")
path([(880, 267), (752, 267)])
label(816, 260, "B, C, D, MRP shares")

# ---------------- funding tags (left) ----------------
TX, TW = 30, 380


def tagrow(row, h, key, vals, hbox=BH):
    y = rows[row] + (hbox - h) / 2
    tag(TX, y, TW, h, key, vals)
    path([(TX + TW, rows[row] + hbox / 2), (SX - 2, rows[row] + hbox / 2)], "e dot", arrow=False)


tagrow("topco", 50, "EQUITY ≈ £238m AT COMPLETION",
       ["Priority £233.9m · A £4.2m · B £0.2m; C, D, MRP classes later"], hbox=62)
tagrow("mid1", 50, "INVESTOR LOAN NOTES · £200m, 12%, DUE 2034",
       ["interest rolls up · TISE: ACTIUM34 · held by BE VII Nominees"])
tagrow("mid2", 36, "SECURITY", ["share charge to GLAS Trust Corp, 16 Aug 2024"])
tagrow("mid3", 36, "SECURITY", ["two charges to GLAS Trust Corp, 16 Aug 2024"])
tagrow("bidco", 50, "SENIOR FACILITIES · AUG 2024",
       ["£230m term loan B (£/$/€, 2031) · £50m acq/capex · £50m RCF"])
tagrow("hold", 50, "TARGET · TAKEN PRIVATE",
       ["scheme effective 19 Aug 2024 · re-registered private 27 Aug 2024"])
tagrow("grp", 36, "GUARANTOR", ["charges to GLAS Trust Corp, 10 Dec 2024"])

# ---------------- side branches (right) ----------------
RX, RW = 800, 340


def side(row, title, num, note, cls="", dashed=False):
    y = rows[row]
    box(RX, y, RW, BH, title, num, note, cls)
    path([(SX + SW, y + BH / 2), (RX - 2, y + BH / 2)], "e dash" if dashed else "e")


side("mid3", "Alpha Alternatives UK Ltd", "09869494", "moved here from AFMC Group, 9 Sep 2024")
side("bidco", "Actium Bidco (US), Ltd.", "Delaware", "parent not on UK register", "unc", dashed=True)
side("hold", "Alpha FMC Trustee Ltd", "10773199", "dormant", "dor")
side("abid", "Alpha FMC Group Nominees Ltd", "08936452", "dormant", "dor")

# ---------------- operating tier ----------------
BW, GAP, OX, OY, OH = 122, 8, 19, 910, 80
cols = [OX + i * (BW + GAP) for i in range(9)]
path([(600, rows["grp"] + BH), (600, 888)], arrow=False)
path([(cols[0] + BW / 2, 888), (cols[7] + BW / 2, 888)], arrow=False)
path([(cols[7] + BW / 2, 888), (cols[8] + BW / 2, 888)], "e dash", arrow=False)
ops = [
    (["Alpha Financial", "Markets Consulting", "UK Ltd"], "04710715", "main UK trading co", "op"),
    (["Alpha (Axxsys)", "Ltd"], "04967647", "acquired 2019", "op"),
    (["Obsidian", "Solutions Ltd"], "09737564", "acquired 2019", "op"),
    (["AIVIQ Ltd"], "11480862", "software", "op"),
    (["Alpha Data", "Solutions Ltd"], "14072313", "dormant", "dor op"),
    (["Alpha FMC", "MENA Ltd"], "14638009", "Middle East", "op"),
    (["White Marble", "Group Ltd"], "13175157", "acquired Sep 2024", "op"),
    (["Alpha (JPSB)", "Group Ltd"], "12195470", "acquired May 2026", "op"),
]
for i, (t, n, note, cls) in enumerate(ops):
    path([(cols[i] + BW / 2, 888), (cols[i] + BW / 2, OY - 2)])
    box(cols[i], OY, BW, OH, t, n, note, cls)

# overseas group box
path([(cols[8] + BW / 2, 888), (cols[8] + BW / 2, OY - 2)], "e dash")
ov = ["USA 5 · Australia 4", "France 3 · Canada 2", "Germany 2", "Switzerland 2", "Singapore 2",
      "Luxembourg", "Netherlands", "Denmark · Hungary", "Hong Kong · Qatar", "Serbia"]
out.append(f'<rect x="{cols[8]}" y="{OY}" width="{BW}" height="226" rx="3" class="bx unc op"/>')
out.append(text_lines(cols[8] + BW / 2, OY + 20,
                      [("t", "27 overseas"), ("t", "subsidiaries")] + [("m", v) for v in ov]
                      + [("m", "parents not on"), ("m", "UK register")]))

# grandchildren
GY1, GY2, GY3, GH = 1012, 1088, 1164, 62
# under AFMC UK: Bankside FS -> Bankside Cloud; Alpha Technology Services (dormant)
c0 = cols[0]
GX, GW = c0 + 16, BW - 16
path([(c0 + 7, OY + OH), (c0 + 7, GY3 + GH / 2)], arrow=False)
path([(c0 + 7, GY1 + GH / 2), (GX - 2, GY1 + GH / 2)])
box(GX, GY1, GW, GH, ["Bankside Financial", "Solutions Consulting"], "11524091", None, "op")
path([(GX + GW / 2, GY1 + GH), (GX + GW / 2, GY2 - 2)])
box(GX, GY2, GW, GH, ["Bankside Cloud", "Technology Ltd"], "15097311", None, "op")
path([(c0 + 7, GY3 + GH / 2), (GX - 2, GY3 + GH / 2)])
box(GX, GY3, GW, GH, ["Alpha Technology", "Services Consulting"], "09244222", "dormant", "dor op")
# White Marble -> WM Consulting ; JPSB Group -> JPSB Ltd
for i, t, n, note in ((6, ["White Marble", "Consulting Ltd"], "09027730", "HSBC charge o/s"),
                      (7, ["Alpha (JPSB) Ltd"], "04156972", "vendor still a PSC")):
    path([(cols[i] + BW / 2, OY + OH), (cols[i] + BW / 2, GY1 - 2)])
    box(cols[i], GY1, BW, 64, t, n, note, "op")

svg = "".join(out)
DEFS = ('<defs><marker id="ah" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="7" markerHeight="7" '
        'orient="auto-start-reverse"><path d="M0,0 L8,4 L0,8 z" class="ahd"/></marker></defs>')

page = open(__file__.replace("org_chart_build.py", "org_chart_template.html")).read()
page = page.replace("{{SVG}}", f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="Ownership chart of the Alpha FMC group: '
                    'Bridgepoint Europe VII funds own Actium Topco (UK) Ltd, which sits above a chain of four 100% holding '
                    'companies (Midco 1, 2, 3 and Bidco) down to the former Alpha FMC plc and the operating subsidiaries. '
                    'Loan notes are issued at Midco 1 and senior debt at Bidco." font-family="\'IBM Plex Sans Condensed\', \'Arial Narrow\', sans-serif">'
                    f'{DEFS}{svg}</svg>')
open(__file__.replace("org_chart_build.py", "org_chart.html"), "w").write(page)
print("written", len(page))
