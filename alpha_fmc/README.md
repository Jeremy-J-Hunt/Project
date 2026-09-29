# Alpha FMC (Bridgepoint) – corporate structure pack

This folder holds the filings used to map the ownership structure of Alpha FMC after Bridgepoint took it private (scheme effective 19 Aug 2024).

## Layout

| Path | Contents |
|---|---|
| `ch_harvest.py` | Pulls Companies House data and PDFs for the whole group and walks up the ownership chain using the PSC (persons with significant control) registers |
| `companies/<number>_<name>/` | For each company: `profile.json`, `officers.json`, `psc.json`, `charges.json` and `filing_history.json` |
| `companies/<number>_<name>/docs/` | Filed PDFs: accounts (all years), plus confirmation statements, SH01 share allotments, MR01 charges, articles and resolutions, incorporation papers, name changes and PSC filings from 2023 onwards |
| `document_index.csv` | One row per document, with its date, type and status |
| `structure.md` | Auto-generated map: non-UK parents, the UK chain, charges and directors |
| `luxembourg/` | Luxembourg filings (RCS extracts, articles, accounts), downloaded manually – see below |
| `deal_documents/` | Scheme document, Rule 2.7 announcement, financing update and TISE listing document |

## Known entities (seeds)

| Company no. | Entity | Role |
|---|---|---|
| 15735478 | Actium Topco (UK) Limited | Top UK holding company (incorporated 22 May 2024) |
| 15735608 | Actium Midco 1 (UK) Limited | Holding company; has securities listed on TISE (The International Stock Exchange) |
| _found by crawl_ | Actium Midco 2 (UK) Limited / Actium Midco 3 (UK) Limited | Bidco's immediate holding companies |
| 15736419 | Actium Bidco (UK) Limited | Acquisition vehicle; borrower of the £220m term loan and £50m revolving credit facility |
| 09965297 | Actium Holdings Limited (formerly Alpha Financial Markets Consulting plc) | Former listed company |
| 07160664 | Alpha Financial Markets Consulting Group Limited | Intermediate holding company |
| 04710715 | Alpha Financial Markets Consulting UK Limited | UK operating company |
| 14638009 | Alpha Financial Markets Consulting MENA Limited | Middle East operating company |

## Running it

```bash
export CH_API_KEY=<key>        # free REST key from developer.company-information.service.gov.uk
python3 alpha_fmc/ch_harvest.py
```

## Manual steps (not reachable through Companies House)

1. **Luxembourg Business Register (lbr.lu):** take the Luxembourg parent(s) named in Actium Topco (UK)'s PSC register and download:
   - the RCS extract
   - the articles (statuts)
   - the filed annual accounts
   - the list of managers

   Trace upwards towards the Bridgepoint Europe VII entities: Bridgepoint Europe VII (GP) S.à r.l. (B257625) and Bridgepoint Europe VII Investments (2) S.à r.l. (B276872).
2. **TISE:** download the listing document for the Actium Midco 1 (UK) securities.
3. **Subsidiaries outside the UK:** the subsidiary note in the plc's last annual report lists them. Cross-check each one against its local register.
