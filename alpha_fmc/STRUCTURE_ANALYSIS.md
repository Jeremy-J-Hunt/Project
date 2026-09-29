# Alpha FMC (Bridgepoint) – corporate structure analysis

_Sources: Companies House filings harvested into `companies/` (38 companies, about 1,340 PDFs), the Actium Topco (UK) Limited group accounts for the year to 31 March 2026, and TISE (The International Stock Exchange) security data. References in [brackets] point to files in this folder. The auto-generated PSC tree is in `structure.md`._

## 1. Summary

- **Sponsor:** Bridgepoint Europe VII (Bridgepoint's flagship buyout fund). The fund's general partner is **Bridgepoint Europe VII GP LLP** (OC438202), an English LLP. Its shares are held by **Bridgepoint Europe VII Nominees Limited** (13793519).
  - **There is no Luxembourg or other offshore holdco** anywhere between the fund and the company.
  - Management and a **Jersey employee benefit trust** (CSC Employee Benefit Trust (Jersey) Limited) sit alongside the fund at Topco.
- **Deal:** a £626m recommended cash offer for Alpha Financial Markets Consulting plc (AIM-listed), made by Actium Bidco (UK) Limited.
  - The scheme became effective on **19 August 2024**.
  - The plc re-registered as a private company and was renamed **Actium Holdings Limited** on 27 August 2024.
- **Stack:** five new UK holding companies, all incorporated 22–23 May 2024:
  - Topco → Midco 1 → Midco 2 → Midco 3 → Bidco
  - Bidco owns the former plc, which sits on top of the legacy trading group.
- **Funding at completion (approximate):**
  - about £238m of sponsor and management equity at Topco
  - a **£200m 12% investor loan note** issued by Midco 1 and listed on TISE
  - a **£230m Facility B term loan** (multi-currency)
  - an undrawn £50m acquisition and capital expenditure facility, plus a £50m RCF
  - GLAS Trust Corporation is security agent.
- **Trading, FY to 31 March 2026:** revenue £366.1m and adjusted EBITDA £66.5m.
  - The FY25 figures (£172.8m revenue) cover only the part-period from 22 May 2024.
  - Acquisitions since the deal: Auxo (US, June 2025, $40m drawn on the acquisition and capex facility), White Marble, Bankside, and JPSB (May 2026, £8.0m).
  - [15735478 FY26 AA]

## 2. Structure chart

```
Bridgepoint Group plc (11443992)
 └ … Bridgepoint Advisers Holdings (01899316)
     ├ Bridgepoint Advisers Ltd (03220373) ─────────────── PSC: significant influence over Topco
     ├ Bridgepoint Europe (SGP) Ltd (SC332267)
     │   └ Bridgepoint Europe VII GP LLP (OC438202) ──── PSC: >75% shares/votes of Topco (as fund GP)
     └ Bridgepoint Europe VII Nominees Ltd (13793519) ─ registered holder of Topco shares and the Midco 1 loan notes
                     │                                   (nominee for the Bridgepoint Europe VII fund partnerships)
                     │   + management shareholders (aggregated in CS01)
                     │   + CSC Employee Benefit Trust (Jersey) Ltd (trustee of the EBT)
                     ▼
ACTIUM TOPCO (UK) LIMITED (15735478)       equity: Priority, A, B, C1/C2, D1/D2, MRP shares
 └ ACTIUM MIDCO 1 (UK) LIMITED (15735608)  issuer: £200m 12% unsecured non-QCB loan notes 2034 (TISE: ACTIUM34)
    └ ACTIUM MIDCO 2 (UK) LIMITED (15735686)  charge to GLAS (16 Aug 2024)
       └ ACTIUM MIDCO 3 (UK) LIMITED (15736277)  2 charges to GLAS (16 Aug 2024)
          ├ ACTIUM BIDCO (UK) LIMITED (15736419)  presumed senior borrower (SFA not public); charge to GLAS
          │  ├ ACTIUM BIDCO (US), LTD. (Delaware)             [per s409 list]
          │  └ ACTIUM HOLDINGS LIMITED (09965297)  ex-Alpha Financial Markets Consulting plc
          │     ├ ALPHA FMC TRUSTEE LIMITED (10773199)       dormant (legacy EBT trustee)
          │     └ ALPHA FMC BIDCO LIMITED (09928343)          legacy 2016 buyout vehicle, kept
          │        ├ ALPHA FMC GROUP NOMINEES LTD (08936452)  dormant
          │        └ ALPHA FINANCIAL MARKETS CONSULTING GROUP LIMITED (07160664)  intermediate holdco; GLAS charges (Dec 2024)
          │           ├ ALPHA FINANCIAL MARKETS CONSULTING UK LIMITED (04710715)  main UK trading company; GLAS charge (Dec 2024)
          │           │  ├ ALPHA TECHNOLOGY SERVICES CONSULTING LIMITED (09244222)  dormant
          │           │  └ BANKSIDE FINANCIAL SOLUTIONS CONSULTING LIMITED (11524091)  acquired Sep 2024
          │           │     └ BANKSIDE CLOUD TECHNOLOGY LTD (15097311)
          │           ├ ALPHA (AXXSYS) LIMITED (04967647)          acquired 2019 (moved under AFMC Group Dec 2020)
          │           ├ OBSIDIAN SOLUTIONS LTD (09737564)          acquired Nov 2019
          │           ├ AIVIQ LIMITED (11480862)
          │           ├ ALPHA DATA SOLUTIONS LIMITED (14072313)    dormant
          │           ├ ALPHA FINANCIAL MARKETS CONSULTING MENA LIMITED (14638009)
          │           ├ WHITE MARBLE GROUP LIMITED (13175157)      acquired Sep 2024
          │           │  └ WHITE MARBLE CONSULTING LIMITED (09027730)  (legacy HSBC charge outstanding)
          │           ├ ALPHA (JPSB) GROUP LIMITED (12195470)      acquired 1 May 2026
          │           │  └ ALPHA (JPSB) LIMITED (04156972)
          │           └ overseas trading subsidiaries (see §5)
          └ ALPHA ALTERNATIVES UK LIMITED (09869494)  moved up from AFMC Group to Midco 3 on 9 Sep 2024
             └ AlphaAlternatives Holdings Inc. (US) / Canada / Australia / US LLC   [per s409 list]
```

The exact parent of each overseas company, and of Bankside Payroll Ltd, is not stated on Companies House. The s409 note says only that every entity except Midco 1 is an *indirect* subsidiary of Topco.

## 3. Entity register (UK)

| Layer | Company | No. | Incorporated | Role | PSC link (date notified) |
|---|---|---|---|---|---|
| Fund | Bridgepoint Europe VII GP LLP | OC438202 | 6 Jul 2021 | GP of Bridgepoint Europe VII | Bridgepoint Europe (SGP) Ltd |
| Fund | Bridgepoint Europe VII Nominees Ltd | 13793519 | 10 Dec 2021 | Legal holder of Topco shares and the loan notes | Bridgepoint Advisers Holdings |
| Fund | Bridgepoint Advisers Ltd | 03220373 | 1996 | Manager; significant influence over Topco | Bridgepoint Advisers Holdings |
| 1 | Actium Topco (UK) Ltd | 15735478 | 22 May 2024 | Group parent; files consolidated accounts | GP LLP and Bridgepoint Advisers (22 May 2024) |
| 2 | Actium Midco 1 (UK) Ltd | 15735608 | 22 May 2024 | Issues the investor loan notes | Topco (22 May 2024) |
| 3 | Actium Midco 2 (UK) Ltd | 15735686 | 22 May 2024 | Holdco; gives share security | Midco 1 (22 May 2024) |
| 4 | Actium Midco 3 (UK) Ltd | 15736277 | 23 May 2024 | Holdco; gives share security; parent of Alpha Alternatives UK | Midco 2 (23 May 2024) |
| 5 | Actium Bidco (UK) Ltd | 15736419 | 23 May 2024 | Offeror; presumed senior borrower | Midco 3 (23 May 2024) |
| 6 | Actium Holdings Ltd (ex-plc) | 09965297 | 22 Jan 2016 | Former listed company (previously Caelius Topco → Alpha FMC Topco → plc) | Bidco (19 Aug 2024) |
| 7 | Alpha FMC Bidco Ltd | 09928343 | 23 Dec 2015 | Legacy holdco (previously Caelius Bidco) | Actium Holdings (18 Oct 2023) |
| 8 | Alpha Financial Markets Consulting Group Ltd | 07160664 | 17 Feb 2010 | Holdco of the trading group | Alpha FMC Bidco (18 Oct 2023) |
| 9 | Alpha Financial Markets Consulting UK Ltd | 04710715 | 25 Mar 2003 | UK trading company | AFMC Group (6 Apr 2016) |

Full details on every company, including officers, charges and filing history, are in `companies/<number>_<name>/*.json`.

## 4. Capital structure

**Equity (Topco).** The initial allotment was on 19 Aug 2024 [SH01 2024-08-23]:
- 233,878,671 **Priority shares** at £1: non-voting, preferred.
- 4,219,800 **A ordinary** at about £1: voting, mostly Bridgepoint.
- 162,763 **B ordinary**: management.
- Later issues: C1/C2 and D1/D2 ordinary (management incentive shares), and MRP ordinary and MRP priority shares (issued 2025–26; the name suggests a management reinvestment plan, but this is unconfirmed).

Total share capital at 31 March 2026 was 249.1m shares. The waterfall in the articles runs as follows:
1. Priority shares and loan notes are repaid pari passu, up to subscription price and principal.
2. Then the accrued priority return and loan-note interest, pari passu.
3. Then A and B ordinary shareholders share the remainder pro rata.

The C and D classes look like management incentive shares; their rights are set out in the articles [MA 2026-06-22].

**Shareholders (CS01, 21 May 2026):**
- Bridgepoint Europe VII Nominees Ltd: 204.1m Priority and 4.07m A shares.
- Management (aggregated): 29.7m Priority, 0.3m B, 0.6m C1/D1, 9.5m MRP Priority and others.
- The 2025 CS01 names individual managers and the **CSC Employee Benefit Trust (Jersey)**.

**Investor loan notes (Midco 1):**
- £200.0m of "12 per cent. unsecured non-QCB loan notes 2034", issued in August 2024 to Bridgepoint Europe VII Nominees Ltd.
- Interest is 12% compounding; the issuer may choose to roll it up rather than pay cash.
- Accrued interest was £40.6m at 31 March 2026. The notes are repayable in March 2034.
- They are listed on TISE (ACTIUM34; listing sponsor Ogier Corporate Finance). Loan notes like these are usually listed to use the quoted-Eurobond exemption from UK withholding tax on interest.
- Sources: [TISE JSON; Midco 1 FY25 AA note 10].

**Senior debt (senior facilities agreement, August 2024; the borrower is presumably Actium Bidco, as the agreement is not public):**
- **Facility B:** £230.0m, due August 2031. After drawdown, £88.0m was redenominated into US dollars ($115.7m) and £44.0m into euros (€52.2m).
- **Acquisition and capex facility:** £50.0m, due August 2031. $40m was drawn in June 2025 to fund the Auxo acquisition.
- **RCF:** £50.0m, SONIA-based, due February 2031. £2.5m was drawn at 31 March 2026.
- **Security and covenant:** there is a net-leverage covenant. Security is held by **GLAS Trust Corporation Limited** as security agent, over shares and material assets of the obligors:
  - interim security on 12 Jul 2024, released on 5 Sep 2024
  - final security on 16 Aug 2024 from Midco 2, Midco 3 and Bidco
  - security from Alpha FMC Group and Alpha FMC UK on 10 Dec 2024 (these companies acceded as guarantors after completion).

**Audit exemption:** the subsidiaries take the s479A audit exemption. Topco files a parent guarantee for each of them (GUARANTEE2/AGREEMENT2 filings).

## 5. Overseas subsidiaries (from the Topco FY26 s409 note)

| Jurisdiction | Entities |
|---|---|
| USA | Actium Bidco (US), Ltd.; Alpha Financial Markets Consulting, Inc.; AlphaAlternatives Holdings Inc.; AlphaAlternatives US LLC; Auxo Solutions LLC; Axxsys Consulting USA, Inc (dormant) |
| Canada | AlphaAlternatives Canada Ltd.; Alpha FM Consulting Canada Inc. |
| Australia | AlphaAlternatives Pty Ltd; Alpha Financial Markets Consulting Australia Pty Ltd; Shoreline Consolidated Pty Ltd; Shoreline Consulting Pty Ltd |
| France | Alpha Financial Markets Consulting S.A.S.; Alpha FMC (Insurance) France S.A.S.; Alpha Technology Services S.A.S. (dormant) |
| Luxembourg | Alpha Financial Markets Consulting (Luxembourg) S.A. (operating subsidiary, Strassen) |
| Germany | Alpha Financial Markets Consulting Germany GmbH; Alpha FMC Alternatives Consulting GmbH |
| Switzerland | Alpha Financial Markets Consulting Switzerland S.A.; Lionpoint Group SA |
| Netherlands / Denmark / Hungary | Alpha FMC Netherlands BV; Alpha FMC Danmark ApS; Alpha FMC Hungary Kft (group services) |
| Hong Kong / Singapore / Qatar / Serbia | Alpha FMC Hong Kong Ltd; Alpha FMC Singapore Pte Ltd; Shoreline Consulting Pte Ltd; Alpha FMC Middle East LLC; Obsidian Alpha Data Solutions LLC Belgrade |

## 6. Legacy structure and history

- **2013–2016:** Alpha FMC Group Holdings Ltd and Alpha FMC Group Ltd (originally EMS Powerstar) were the earlier holding companies. The 2013 facilities were from Baird Capital and Beechbrook Mezzanine.
- **Dec 2015 – Jan 2016:** the "Caelius" buyout stack was set up: Caelius Topco → Caelius Midco → Midco 2 → Caelius Bidco, with Lloyds senior debt and Beechbrook mezzanine. Caelius Topco became the plc and listed on AIM in 2017.
- **Oct–Dec 2023:** before the Bridgepoint deal, the legacy intermediate holdcos were taken out of the chain:
  - Alpha FMC Bidco was moved directly under the plc (18 Oct 2023).
  - Alpha FMC Midco, Midco 2, Group Holdings and Group Ltd went into members' voluntary liquidation (solvent winding-up) on 14 Dec 2023 and were dissolved on 14 Feb 2025.
- **19 Aug 2024:** the Bridgepoint scheme became effective.
- **27 Aug 2024:** the plc re-registered as a private company and was renamed Actium Holdings Ltd.
- **Sep 2024:** White Marble Group and Bankside Financial Solutions joined the group, and Alpha Alternatives UK was moved under Midco 3.

## 7. Points to follow up

1. **Stale charges.** Alpha FMC Group Ltd (07160664) still shows two **9 March 2016 charges as Outstanding**: Lloyds (senior) and Beechbrook (mezzanine). That debt was almost certainly repaid at the 2017 IPO, so a satisfaction filing (MR04) appears to be missing. White Marble Consulting's 2020 HSBC debenture is also still outstanding.
2. **Bankside Payroll Ltd (12406540)** has only a PSC *statement* on its register, and its individual PSCs ceased on 1 Jan 2024. It is listed as a group subsidiary in the s409 note, but no corporate parent is registered. Check its CS01 shareholder list.
3. **Alpha (JPSB) Ltd** still lists Kirandeep Singh Bhogal (the vendor) as a PSC with *significant influence* alongside JPSB Group. This may reflect earn-out or vendor rights, or a filing that hasn't been updated since completion in May 2026.
4. **Why Alpha Alternatives UK sits under Midco 3** (from 9 Sep 2024) rather than under the operating group. It may relate to how the US assets were put up as security, or to the Actium Bidco (US) debt tranche. The SFA isn't public, so check the MR01s and the FY25 accounts of Midco 3.
5. **Loan-note issuer wording.** The Topco consolidated accounts say the notes were "issued by the Company". The Midco 1 accounts and TISE both show **Midco 1** as issuer. Treat Midco 1 as correct.
6. **Overseas registries.** The US (Delaware), Swiss, Australian and other subsidiaries have not been pulled. Use the s409 addresses in §5 to look them up in each local registry.

## 8. Folder guide

| Path | Contents |
|---|---|
| `STRUCTURE_ANALYSIS.md` | This document |
| `structure.md` | Auto-generated PSC tree, charges, directors and latest accounts (`python3 ch_harvest.py --rebuild`) |
| `document_index.csv` | One row per PDF: company, date, filing type, PDF path and OCR text path |
| `companies/<no>_<name>/` | `profile`, `officers`, `psc`, `charges` and `filing_history` JSON, plus `docs/` (PDFs, with a `.txt` OCR copy next to each) |
| `deal_documents/` | TISE security record for the ACTIUM34 loan notes |
| `ch_harvest.py` / `ocr_docs.py` | Re-run to refresh the data or OCR new PDFs |

**Key documents to read first:**
- Topco group accounts FY26 and FY25: `companies/15735478_*/docs/2026-07-24_AA_*` and `2025-08-12_AA_*`
- Topco confirmation statements (CS01), 2025 and 2026
- Topco first share allotment (SH01, 2024-08-23)
- Topco articles (MA, 2026-06-22)
- Midco 1 accounts (AA, 2026-01-08)
- GLAS charge filings (MR01) under Midco 2, Midco 3 and Bidco
- Actium Holdings re-registration filings, 27 Aug 2024 (CERT11 / RR02)
