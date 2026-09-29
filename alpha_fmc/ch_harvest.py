#!/usr/bin/env python3
"""Harvest Companies House data for the Alpha FMC / Actium (Bridgepoint) group.

Scrapes the public Companies House website (no API key needed): profile,
officers, PSCs, charges and filing history for every group entity, downloads
the filed PDFs, walks up the ownership chain via corporate PSCs, and writes a
structure map.

Usage:
    python3 alpha_fmc/ch_harvest.py            # full run
    python3 alpha_fmc/ch_harvest.py --no-docs  # metadata only
    python3 alpha_fmc/ch_harvest.py --extra 12345678 SC123456   # add companies
"""
import argparse
import csv
import html
import json
import re
import time
from pathlib import Path

import requests

BASE = "https://find-and-update.company-information.service.gov.uk"
ROOT = Path(__file__).resolve().parent
OUT = ROOT / "companies"

# Known group entities. The crawl adds more via PSCs and name searches.
SEEDS = [
    "15735478",  # Actium Topco (UK) Limited
    "15735608",  # Actium Midco 1 (UK) Limited
    "15736419",  # Actium Bidco (UK) Limited
    "09965297",  # Actium Holdings Limited (formerly Alpha Financial Markets Consulting plc)
    "07160664",  # Alpha Financial Markets Consulting Group Limited
    "04710715",  # Alpha Financial Markets Consulting UK Limited
    "14638009",  # Alpha Financial Markets Consulting MENA Limited
    # UK subsidiaries listed in the Topco FY26 s409 related-undertakings note
    "11480862",  # AIVIQ Limited
    "09869494",  # Alpha Alternatives UK Limited
    "14072313",  # Alpha Data Solutions Limited
    "09244222",  # Alpha Technology Services Consulting Limited
    "04967647",  # Alpha (Axxsys) Limited
    "15097311",  # Bankside Cloud Technology Ltd
    "11524091",  # Bankside Financial Solutions Consulting Limited
    "12406540",  # Bankside Payroll Limited
    "09737564",  # Obsidian Solutions Ltd
    "09027730",  # White Marble Consulting Limited
    "13175157",  # White Marble Group Limited
    "12195470",  # Alpha (JPSB) Group Limited (acquired May 2026)
    "04156972",  # Alpha (JPSB) Limited
]
# Name-search hits that are not group companies.
EXCLUDE = {"06375145", "08394215", "02887088"}
SEARCH_TERMS = ["actium", "alpha financial markets consulting", "alpha fmc"]
NAME_RE = re.compile(r"^actium (topco|midco|bidco|holdings)|alpha (financial markets|fmc)", re.I)
# Only follow PSC links to these (the sponsor side), so the crawl doesn't wander
# into every Bridgepoint portfolio company.
FOLLOW_PSC_RE = re.compile(r"actium|alpha|bridgepoint", re.I)

# Filing types not worth downloading: officer appointments/changes, address moves.
SKIP_TYPES = re.compile(r"^(AP0\d|AP0\dA|TM0\d|CH0\d|AD0\d|AD02|AD03|AD04|PSC04)$")

SESSION = requests.Session()
SESSION.headers["User-Agent"] = "Mozilla/5.0 (research; structure mapping)"


def get(path, **params):
    for attempt in range(6):
        r = SESSION.get(BASE + path, params=params, timeout=60)
        if r.status_code == 429 or r.status_code >= 500:
            time.sleep(10 * (attempt + 1))
            continue
        if r.status_code == 404:
            return None
        r.raise_for_status()
        time.sleep(0.4)
        return r.text
    raise RuntimeError(f"failed: {path}")


def text(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s or ""))).strip()


def slug(s, n=60):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")[:n]


def dl_pairs(block):
    """<dt>label</dt><dd>value</dd> pairs -> dict."""
    return {text(k): text(v) for k, v in re.findall(r"<dt[^>]*>(.*?)</dt>\s*<dd[^>]*>(.*?)</dd>", block, re.S)}


def parse_profile(number):
    t = get(f"/company/{number}")
    if not t:
        return None
    name = text(re.search(r'<p class="heading-xlarge"[^>]*>(.*?)</p>|<h1[^>]*>(.*?)</h1>', t, re.S).group(0))
    body = t[t.find('id="page-container"'):]
    prof = {"company_number": number, "company_name": name}
    prof.update({k: v for k, v in dl_pairs(body).items() if len(k) < 60})
    sic = re.findall(r'id="sic\d+"[^>]*>(.*?)</span>', t, re.S)
    prof["sic"] = [text(s) for s in sic]
    prev = re.findall(r'id="previous-name-\d+"[^>]*>(.*?)</td>', t, re.S)
    prof["previous_names"] = [text(p) for p in prev]
    return prof


def split_blocks(t, cls):
    parts = re.split(rf'<div class="{cls}-\d+"', t)
    return parts[1:]


def parse_officers(number):
    out, page = [], 1
    while True:
        t = get(f"/company/{number}/officers", page=page)
        if not t:
            break
        blocks = split_blocks(t, "appointment")
        for b in blocks:
            o = {"name": text(re.search(r'id="officer-name-\d+">(.*?)</span>', b, re.S).group(1))}
            o.update(dl_pairs(b))
            st = re.search(r'officer-status-tag-\d+"[^>]*>(.*?)</span>', b, re.S)
            o["status"] = text(st.group(1)) if st else ""
            out.append(o)
        if f"page={page + 1}" not in t or not blocks:
            break
        page += 1
    return out


def parse_psc(number):
    t = get(f"/company/{number}/persons-with-significant-control")
    out = []
    if not t:
        return out
    for b in re.split(r'<div class="appointment-\d+"', t)[1:]:
        nm = re.search(r'id="psc-name-\d+"[^>]*>(.*?)</(?:span|b|strong|h2)>', b, re.S) or \
            re.search(r"<h2[^>]*>(.*?)</h2>", b, re.S)
        p = {"name": text(nm.group(1)) if nm else ""}
        p.update(dl_pairs(b))
        st = re.search(r'status-tag[^"]*"[^>]*>(.*?)</span>', b, re.S)
        p["status"] = text(st.group(1)) if st else ""
        noc = re.findall(r'id="psc-noc-\d+[^"]*"[^>]*>(.*?)</', b, re.S) or \
            re.findall(r"<li[^>]*>(.*?)</li>", b.split("Nature of control")[-1], re.S)
        p["natures_of_control"] = [text(n) for n in noc if text(n)]
        out.append(p)
    return out


def parse_charges(number):
    out, page = [], 1
    while True:
        t = get(f"/company/{number}/charges", page=page)
        if not t:
            break
        blocks = split_blocks(t, "mortgage")
        for b in blocks:
            c = {"charge": text(re.search(r'id="mortgage-heading-\d+"[^>]*>(.*?)</a>', b, re.S).group(1))}
            c.update(dl_pairs(b))
            c["persons_entitled"] = [text(x) for x in re.findall(r'id="persons-entitled-\d+-charge-\d+"[^>]*>(.*?)</li>', b, re.S)]
            d = re.search(r'id="mortgage-particulars-\d+"[^>]*>(.*?)</p>', b, re.S)
            c["description"] = text(d.group(1)) if d else ""
            out.append(c)
        if f"page={page + 1}" not in t or not blocks:
            break
        page += 1
    return out


def parse_filings(number):
    out, page = [], 1
    while True:
        t = get(f"/company/{number}/filing-history", page=page)
        if not t:
            break
        rows = re.findall(r"<tr>(.*?)</tr>", t, re.S)
        got = 0
        for r in rows:
            tds = re.findall(r"<td[^>]*>(.*?)</td>", r, re.S)
            if len(tds) < 3:
                continue
            got += 1
            link = re.search(r'href="(/company/[^"]+/document\?format=pdf[^"]*)"', r)
            pages = re.search(r"\((\d+) pages?\)", r)
            out.append({"date": text(tds[0]), "type": text(tds[1]),
                        "description": text(tds[2]),
                        "pdf": html.unescape(link.group(1)) if link else None,
                        "pages": int(pages.group(1)) if pages else None})
        if not got or f"page={page + 1}" not in t:
            break
        page += 1
    return out


def iso(d):
    return time.strftime("%Y-%m-%d", time.strptime(d, "%d %b %Y")) if re.match(r"\d+ \w{3} \d{4}$", d) else d


def search(term):
    t = get("/search/companies", q=term) or ""
    hits = re.findall(r'<a class="govuk-link" href="/company/([A-Z0-9]+)"[^>]*>(.*?)</a>', t, re.S)
    return [(n, text(nm)) for n, nm in hits]


def psc_company_number(p):
    """UK company/LLP number for a corporate PSC registered in the UK, else None."""
    reg = p.get("Registration number", "").upper().replace(" ", "")
    where = (p.get("Place registered", "") + " " + p.get("Incorporated in", "")).lower()
    if reg and re.search(r"england|wales|united kingdom|scotland|companies house", where) \
            and re.fullmatch(r"(SC|NI|OC|SO)?\d{6,8}", reg):
        return reg if not reg.isdigit() else reg.zfill(8)
    return None


def light_filter(filings):
    """Sponsor-side entities: last 3 accounts and last 2 confirmation statements only."""
    aa = [f for f in filings if f["type"].startswith("AA")][:3]
    cs = [f for f in filings if f["type"].startswith("CS01")][:2]
    return aa + cs


def harvest(number, docs, light=False):
    prof = parse_profile(number)
    if not prof:
        print(f"  ! {number} not found")
        return None
    name = prof["company_name"]
    d = OUT / f"{number}_{slug(name)}"
    (d / "docs").mkdir(parents=True, exist_ok=True)
    data = {"profile": prof, "officers": parse_officers(number), "psc": parse_psc(number),
            "charges": parse_charges(number), "filing_history": parse_filings(number)}
    for k, v in data.items():
        (d / f"{k}.json").write_text(json.dumps(v, indent=2, ensure_ascii=False))

    rows, seen = [], {}
    for f in (light_filter(data["filing_history"]) if light else data["filing_history"]):
        if not f["pdf"] or SKIP_TYPES.match(f["type"]):
            continue
        ftype = re.sub(r"[^A-Za-z0-9]+", "-", f["type"])  # e.g. MEM/ARTS
        stem = f"{iso(f['date'])}_{ftype}_{slug(f['description'], 50)}"
        seen[stem] = seen.get(stem, 0) + 1  # same-day filings of the same type
        fname = f"{stem}.pdf" if seen[stem] == 1 else f"{stem}-{seen[stem]}.pdf"
        dest = d / "docs" / fname
        status = "exists" if dest.exists() else "skipped"
        if docs and not dest.exists():
            try:
                r = SESSION.get(BASE + f["pdf"], timeout=180)
                r.raise_for_status()
                if not r.content.startswith(b"%PDF"):
                    raise ValueError("not a PDF")
                dest.write_bytes(r.content)
                status = "downloaded"
                time.sleep(0.4)
            except Exception as e:
                status = f"failed: {e}"
        rows.append({"company_number": number, "company_name": name, "date": iso(f["date"]),
                     "type": f["type"], "description": f["description"], "pages": f["pages"],
                     "file": str(dest.relative_to(ROOT)), "status": status})
    print(f"  {number} {name}: {len(data['filing_history'])} filings, {len(rows)} docs, "
          f"{len(data['psc'])} PSCs, {len(data['charges'])} charges")
    return {"number": number, "name": name, "data": data, "docs": rows}


def short_noc(natures):
    """Compress PSC natures of control for the tree view."""
    out = []
    for n in natures:
        n = n.lower()
        if "ownership of shares" in n:
            out.append("shares " + re.sub(r".*shares\s*[–-]\s*", "", n))
        elif "significant influence" in n:
            out.append("sig. influence")
        elif "surplus assets" in n:
            out.append("LLP surplus assets")
    return ", ".join(dict.fromkeys(out)) or "control"


def build_map(results):
    by_num = {r["number"]: r for r in results}
    status = {n: r["data"]["profile"].get("Company status", "") for n, r in by_num.items()}
    edges, tops, statements = [], [], []
    for r in results:
        for p in r["data"]["psc"]:
            if p.get("status", "").lower() in ("ceased", "withdrawn") or p.get("Ceased on"):
                continue
            if p["name"].startswith("Statement"):
                statements.append(r["number"])
                continue
            parent = psc_company_number(p)
            if parent and parent in by_num:
                edges.append((parent, r["number"], short_noc(p.get("natures_of_control", []))))
            else:
                tops.append((p["name"], p.get("Legal form", "") or "individual",
                             p.get("Place registered", "") or p.get("Country of residence", ""),
                             p.get("Registration number", ""), r["number"],
                             "; ".join(p.get("natures_of_control", []))))

    children = {}
    for parent, child, nat in edges:
        children.setdefault(parent, []).append((child, nat))
    has_parent = {c for _, c, _ in edges}
    # Where a company has several corporate controllers, expand it under the one that holds
    # the shares (the others are shown as cross-references).
    primary = {}
    for parent, child, nat in sorted(edges, key=lambda e: "shares" not in e[2] and "surplus" not in e[2]):
        primary.setdefault(child, parent)

    def label(n):
        st = status[n]
        return f"{by_num[n]['name']} [{n}]" + (f"  ({st.upper()})" if st and st != "Active" else "")

    L = ["# Alpha FMC / Actium – Companies House structure map", "",
         "_Generated by `ch_harvest.py` from the PSC registers. Each arrow is a registrable "
         "PSC/RLE link (the nature of control is in brackets). The PSC register is not a shareholder "
         "list, so check it against the CS01 and SH01 filings in each `docs/` folder. Companies with "
         "more than one corporate controller are expanded under the one holding the shares; other links are marked ↔._", "",
         "## Ownership tree", "", "```"]
    printed = set()

    def walk(n, depth, nat, parent=None):
        prefix = "    " * depth + ("└─ " if depth else "")
        tag = f"  ‹{nat}›" if nat else ""
        if n in printed or (parent and primary.get(n) != parent):
            L.append(f"{prefix}{label(n)}{tag}  ↔ shown under {by_num[primary[n]]['name']}")
            return
        printed.add(n)
        L.append(f"{prefix}{label(n)}{tag}")
        for c, cn in sorted(children.get(n, []), key=lambda x: by_num[x[0]]["name"]):
            walk(c, depth + 1, cn, n)

    for root in sorted(set(by_num) - has_parent, key=lambda n: (n not in children, by_num[n]["name"])):
        walk(root, 0, "")
    L += ["```", ""]
    if statements:
        L += ["Companies whose PSC register holds a **statement** instead of (or as well as) a named "
              "controller: " + ", ".join(f"{by_num[n]['name']} [{n}]" for n in sorted(set(statements))), ""]
    L += ["## Individuals / other controllers named on PSC registers", "",
          "| Controller | Legal form | Registered / residence | Reg. no. | Controls | Nature of control |",
          "|---|---|---|---|---|---|"]
    for nm, form, where, reg, child, nat in tops:
        L.append(f"| {nm} | {form} | {where} | {reg} | {by_num[child]['name']} ({child}) | {nat} |")
    L += ["", "## Charges (outstanding first)", "",
          "| Company | Charge | Created | Status | Persons entitled | Description |", "|---|---|---|---|---|---|"]
    rows = [(r, c) for r in results for c in r["data"]["charges"]
            if not r["name"].startswith("BRIDGEPOINT")]
    rows.sort(key=lambda rc: (not rc[1].get("Status", "").startswith("Outstanding"), rc[0]["name"]))
    for r, c in rows:
        L.append(f"| {r['name']} | {c['charge']} | {c.get('Created', '')} | {c.get('Status', '')} | "
                 f"{'; '.join(c['persons_entitled'])} | {c['description'][:100]} |")
    L += ["", "_Charges registered against Bridgepoint entities (fund-level facilities) are left out; "
          "they are in each company's `charges.json`._", "", "## Current directors / officers", ""]
    for r in results:
        cur = [f"{o['name']} ({(o.get('Role') or o.get('Role Active') or '').strip()})"
               for o in r["data"]["officers"] if o.get("status", "").lower() == "active"]
        L.append(f"- **{r['name']}** [{r['number']}]: {'; '.join(cur) or '–'}")
    L += ["", "## Latest accounts filed", ""]
    for r in results:
        aa = [f for f in r["data"]["filing_history"] if f["type"].startswith("AA")]
        if aa:
            L.append(f"- **{r['name']}**: {aa[0]['date']} – {aa[0]['description']}")
    (ROOT / "structure.md").write_text("\n".join(L) + "\n")

def load_saved():
    """Reload every harvested company from companies/*/ (no network)."""
    results = []
    for d in sorted(OUT.iterdir()):
        if not (d / "profile.json").exists():
            continue
        data = {k: json.loads((d / f"{k}.json").read_text())
                for k in ("profile", "officers", "psc", "charges", "filing_history")}
        number, name = data["profile"]["company_number"], data["profile"]["company_name"]
        rows = []
        for pdf in sorted((d / "docs").glob("*.pdf")):
            date, ftype = pdf.name.split("_")[:2]
            rows.append({"company_number": number, "company_name": name, "date": date, "type": ftype,
                         "file": str(pdf.relative_to(ROOT)),
                         "text": str(pdf.with_suffix(".txt").relative_to(ROOT)) if pdf.with_suffix(".txt").exists() else ""})
        results.append({"number": number, "name": name, "data": data, "docs": rows})
    return results


def write_outputs(results):
    rows = [row for r in results for row in r["docs"]]
    with open(ROOT / "document_index.csv", "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]) if rows else ["file"])
        w.writeheader()
        w.writerows(rows)
    build_map(results)
    print(f"\n{len(results)} companies, {len(rows)} documents -> {OUT}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-docs", action="store_true")
    ap.add_argument("--extra", nargs="*", default=[])
    ap.add_argument("--rebuild", action="store_true",
                    help="rebuild structure.md / document_index.csv from saved JSON, no network")
    args = ap.parse_args()

    if args.rebuild:
        results = load_saved()
        write_outputs(results)
        return

    queue = SEEDS + args.extra
    for term in SEARCH_TERMS:
        for n, nm in search(term):
            if NAME_RE.search(nm) and n not in queue and n not in EXCLUDE:
                queue.append(n)

    group = set(queue)
    done, results = set(), []
    while queue:
        n = queue.pop(0)
        if n in done:
            continue
        done.add(n)
        r = harvest(n, docs=not args.no_docs, light=n not in group)
        if not r:
            continue
        results.append(r)
        for p in r["data"]["psc"]:
            parent = psc_company_number(p)
            ceased = p.get("status", "").lower() == "ceased" or p.get("Ceased on")
            if parent and not ceased and parent not in done and FOLLOW_PSC_RE.search(p["name"]):
                queue.append(parent)

    write_outputs(results)


if __name__ == "__main__":
    main()
