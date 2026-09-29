# Alpha FMC (Bridgepoint) – corporate structure pack

Companies House filings and structure mapping for Alpha FMC, which Bridgepoint Europe VII took private (scheme effective 19 Aug 2024).

**Start with [`STRUCTURE_ANALYSIS.md`](STRUCTURE_ANALYSIS.md)**. It covers the structure chart, the capital structure, the legacy history and open points.

| Path | Contents |
|---|---|
| `STRUCTURE_ANALYSIS.md` | Written analysis of the structure (hand-written, with sources cited) |
| `structure.md` | Auto-generated ownership tree from the PSC registers, plus charges, directors and latest accounts |
| `document_index.csv` | One row per PDF: company, date, filing type, PDF path and OCR text path |
| `companies/<number>_<name>/` | `profile.json`, `officers.json`, `psc.json`, `charges.json` and `filing_history.json` |
| `companies/<number>_<name>/docs/` | Every filed PDF except routine director and address changes, with a `.txt` OCR copy next to each |
| `deal_documents/` | TISE record for the £200m 12% loan notes issued by Actium Midco 1 (UK) Ltd (ACTIUM34) |
| `ch_harvest.py` | Scrapes the public Companies House site (no API key needed) and walks the PSC chain upwards |
| `ocr_docs.py` | Adds a text layer (`.txt`) for every PDF, using OCR where the PDF is scanned (needs `tesseract-ocr` and `poppler-utils`) |

**Coverage:** 38 companies.
- The whole UK group: the Actium stack, the former plc, legacy holdcos and the UK operating subsidiaries.
- The Bridgepoint sponsor chain up to Bridgepoint Group plc. For these companies only the recent accounts and confirmation statements are downloaded.

```bash
python3 alpha_fmc/ch_harvest.py            # refresh everything (skips PDFs already downloaded)
python3 alpha_fmc/ch_harvest.py --rebuild  # regenerate structure.md / document_index.csv offline
python3 alpha_fmc/ocr_docs.py              # OCR any PDFs without a .txt
```

**Not covered:** overseas subsidiaries (US, Canada, Australia, France, Luxembourg, Switzerland and others). They're listed in §5 of the analysis with their registered addresses, and need to be looked up in each local registry.
