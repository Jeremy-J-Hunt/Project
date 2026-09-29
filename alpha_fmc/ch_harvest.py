#!/usr/bin/env python3
"""Harvest Companies House data for the Alpha FMC / Actium (Bridgepoint) group.

Pulls profile, officers, PSCs, charges and filing history for every entity in
the group, downloads the key filed documents as PDFs, walks up the ownership
chain via corporate PSCs, and writes a structure map.

Usage:
    export CH_API_KEY=...   # free key: https://developer.company-information.service.gov.uk/
    python3 alpha_fmc/ch_harvest.py            # full run
    python3 alpha_fmc/ch_harvest.py --no-docs  # metadata only

Network hosts needed: api.company-information.service.gov.uk,
document-api.company-information.service.gov.uk and the S3 host that document
downloads redirect to (*.amazonaws.com).
"""
import argparse
import csv
import json
import os
import re
import sys
import time
from pathlib import Path

import requests

API = "https://api.company-information.service.gov.uk"
DOC_API = "https://document-api.company-information.service.gov.uk"
ROOT = Path(__file__).resolve().parent
OUT = ROOT / "companies"

# Known group entities (company number -> note). The crawl adds more via PSCs
# and name searches.
SEEDS = {
    "15735478": "Actium Topco (UK) Limited - top UK holdco",
    "15735608": "Actium Midco 1 (UK) Limited - issuer of TISE-listed notes",
    "15736419": "Actium Bidco (UK) Limited - acquisition vehicle",
    "09965297": "Actium Holdings Limited (formerly Alpha Financial Markets Consulting plc)",
    "07160664": "Alpha Financial Markets Consulting Group Limited",
    "04710715": "Alpha Financial Markets Consulting UK Limited",
    "14638009": "Alpha Financial Markets Consulting MENA Limited",
}
SEARCH_TERMS = ["actium", "alpha financial markets consulting", "alpha fmc"]
# Name filter so searches don't pull in unrelated companies.
NAME_RE = re.compile(r"\bactium\b|alpha (financial markets|fmc)", re.I)

# Filing categories to download as PDFs. Accounts are pulled for all years;
# everything else only from DOCS_SINCE onwards (the deal was 2024).
DOC_CATEGORIES = {
    "accounts", "confirmation-statement", "capital", "mortgage", "incorporation",
    "resolution", "change-of-name", "persons-with-significant-control",
    "reregistration", "insolvency",
}
DOCS_SINCE = "2023-01-01"


class CH:
    def __init__(self, key):
        self.s = requests.Session()
        self.s.auth = (key, "")

    def get(self, path, **params):
        for attempt in range(5):
            r = self.s.get(API + path, params=params, timeout=60)
            if r.status_code == 429:
                time.sleep(60)
                continue
            if r.status_code == 404:
                return None
            r.raise_for_status()
            time.sleep(0.6)  # stay under 600 requests / 5 min
            return r.json()
        raise RuntimeError(f"rate limited on {path}")

    def paged(self, path, key="items", page=100, cap=5000):
        items, start = [], 0
        while start < cap:
            d = self.get(path, items_per_page=page, start_index=start)
            if not d:
                break
            batch = d.get(key, [])
            items += batch
            start += len(batch)
            if not batch or start >= d.get("total_count", d.get("total_results", 0)):
                break
        return items

    def download(self, meta_url, dest):
        # Metadata links may point at frontend-doc-api; the content endpoint is on document-api.
        doc_id = meta_url.rstrip("/").split("/")[-1]
        r = self.s.get(f"{DOC_API}/document/{doc_id}/content", headers={"Accept": "application/pdf"},
                       timeout=120, allow_redirects=False)
        if r.status_code in (301, 302, 303, 307):
            r = requests.get(r.headers["Location"], timeout=120)  # pre-signed S3, no auth
        r.raise_for_status()
        dest.write_bytes(r.content)
        time.sleep(0.6)


def slug(s, n=60):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")[:n]


def uk_number(psc):
    """Return a UK company number for a corporate PSC registered at Companies House."""
    ident = psc.get("identification") or {}
    reg = (ident.get("registration_number") or "").strip().upper().replace(" ", "")
    where = " ".join(str(ident.get(k, "")) for k in
                     ("place_registered", "country_registered", "legal_authority")).lower()
    uk = any(w in where for w in ("companies house", "england", "wales", "united kingdom", "scotland"))
    if reg and uk and re.fullmatch(r"(SC|NI|OC|SO)?\d{6,8}", reg):
        return reg.zfill(8) if reg.isdigit() else reg
    return None


def harvest(ch, number, docs):
    prof = ch.get(f"/company/{number}")
    if not prof:
        print(f"  ! {number} not found")
        return None
    name = prof["company_name"]
    d = OUT / f"{number}_{slug(name)}"
    (d / "docs").mkdir(parents=True, exist_ok=True)
    data = {
        "profile": prof,
        "officers": ch.paged(f"/company/{number}/officers"),
        "psc": ch.paged(f"/company/{number}/persons-with-significant-control"),
        "charges": ch.paged(f"/company/{number}/charges"),
        "filing_history": ch.paged(f"/company/{number}/filing-history"),
    }
    for k, v in data.items():
        (d / f"{k}.json").write_text(json.dumps(v, indent=2))

    rows = []
    for f in data["filing_history"]:
        cat, date = f.get("category", ""), f.get("date", "")
        meta = (f.get("links") or {}).get("document_metadata")
        if cat not in DOC_CATEGORIES or not meta:
            continue
        if cat != "accounts" and date < DOCS_SINCE:
            continue
        fname = f"{date}_{f.get('type', '')}_{slug(f.get('description', ''), 40)}.pdf"
        dest = d / "docs" / fname
        status = "exists"
        if docs and not dest.exists():
            try:
                ch.download(meta, dest)
                status = "downloaded"
            except Exception as e:  # keep going; record the failure
                status = f"failed: {e}"
        elif not docs:
            status = "skipped"
        rows.append({"company_number": number, "company_name": name, "date": date,
                     "category": cat, "type": f.get("type", ""),
                     "description": f.get("description", ""),
                     "file": str(dest.relative_to(ROOT)), "status": status})
    print(f"  {number} {name}: {len(data['filing_history'])} filings, {len(rows)} docs")
    return {"number": number, "name": name, "dir": d, "data": data, "docs": rows}


def build_map(results):
    """Write structure.md: ownership edges from PSC registers, plus charges."""
    by_num = {r["number"]: r for r in results}
    edges, foreign = [], []
    for r in results:
        for p in r["data"]["psc"]:
            if p.get("ceased_on"):
                continue
            nat = ", ".join(p.get("natures_of_control", []))
            parent = uk_number(p)
            label = p.get("name", "?")
            if parent:
                edges.append((parent, r["number"], nat))
            else:
                ident = p.get("identification") or {}
                where = ident.get("country_registered") or ident.get("place_registered") or p.get("kind")
                foreign.append((label, where, ident.get("registration_number", ""), r["number"], nat))

    children = {}
    for parent, child, _ in edges:
        children.setdefault(parent, []).append(child)
    has_parent = {c for _, c, _ in edges}
    lines = ["# Alpha FMC / Actium - Companies House structure map", "",
             "_Generated by ch_harvest.py from PSC registers. Ownership shown is the "
             "registrable relevant legal entity (RLE) for each company._", "",
             "## Non-UK / non-CH controllers (top of chain)", ""]
    for label, where, reg, child, nat in foreign:
        cname = by_num.get(child, {}).get("name", child)
        lines.append(f"- **{label}** ({where}{', reg ' + reg if reg else ''}) → {cname} ({child}) - {nat}")
    lines += ["", "## UK chain", "", "```"]

    def walk(n, depth, seen):
        nm = by_num.get(n, {}).get("name", "(not harvested)")
        lines.append("    " * depth + f"{nm} [{n}]")
        for c in sorted(children.get(n, [])):
            if c not in seen:
                walk(c, depth + 1, seen | {c})

    for root in sorted(set(by_num) - has_parent):
        walk(root, 0, {root})
    lines += ["```", "", "## Charges (security registered)", "",
              "| Company | Created | Status | Persons entitled | Description |",
              "|---|---|---|---|---|"]
    for r in results:
        for c in r["data"]["charges"]:
            ents = "; ".join(p.get("name", "") for p in c.get("persons_entitled", []))
            desc = (c.get("particulars") or {}).get("description", "") or c.get("classification", {}).get("description", "")
            lines.append(f"| {r['name']} | {c.get('created_on', '')} | {c.get('status', '')} | {ents} | {desc[:120]} |")
    lines += ["", "## Directors (current)", ""]
    for r in results:
        cur = [o["name"] for o in r["data"]["officers"] if not o.get("resigned_on")]
        lines.append(f"- **{r['name']}**: {'; '.join(cur)}")
    (ROOT / "structure.md").write_text("\n".join(lines) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-docs", action="store_true", help="metadata only, skip PDFs")
    args = ap.parse_args()
    key = os.environ.get("CH_API_KEY")
    if not key:
        sys.exit("Set CH_API_KEY (free from https://developer.company-information.service.gov.uk/)")
    ch = CH(key)

    queue = list(SEEDS)
    for term in SEARCH_TERMS:
        for it in ch.paged("/search/companies?q=" + requests.utils.quote(term), cap=300):
            if NAME_RE.search(it.get("title", "")) and it["company_number"] not in queue:
                queue.append(it["company_number"])

    done, results = set(), []
    while queue:
        n = queue.pop(0)
        if n in done:
            continue
        done.add(n)
        r = harvest(ch, n, docs=not args.no_docs)
        if not r:
            continue
        results.append(r)
        for p in r["data"]["psc"]:  # walk up the chain
            parent = uk_number(p)
            if parent and parent not in done:
                queue.append(parent)

    with open(ROOT / "document_index.csv", "w", newline="") as fh:
        rows = [row for r in results for row in r["docs"]]
        w = csv.DictWriter(fh, fieldnames=list(rows[0]) if rows else ["file"])
        w.writeheader()
        w.writerows(rows)
    build_map(results)
    print(f"\n{len(results)} companies harvested -> {OUT}\nMap: {ROOT/'structure.md'}")


if __name__ == "__main__":
    main()
