# Clinuvel Pharmaceuticals — could it be a private-equity target?

A buy-side deep dive on Clinuvel Pharmaceuticals (ASX: CUV, Nasdaq: CUVL) as of 22 September 2026, with a working take-private LBO model ("Project Daylight").

**Read first: [`Clinuvel_PE_Deep_Dive.md`](Clinuvel_PE_Deep_Dive.md)**

## Verdict in brief

- **It looks like a PE target.**
  - Ten consecutive profitable years and A$252m of cash with no debt.
  - The share price (about A$8.24) is only about 1.7x cash per share. That puts enterprise value (EV) at about 4x EBITDA.
  - There is about A$20–25m of discretionary cost to cut.
  - Shareholders are disaffected: there have been three remuneration-report "strikes".
- **It is a poor LBO.**
  - SCENESSE is effectively the only product. Its monopoly in EPP (erythropoietic protoporphyria, a rare light-sensitivity disorder) faces two oral challengers within about nine months:
    - LEO Pharma's dersimelagon, with an FDA decision expected by about February 2027;
    - Disc Medicine's bitopertin, with Phase 3 results due in Q4 2026 and a decision by mid-2027.
  - Orphan-drug exclusivity is ending in the US and has already ended in the EU.
  - US sales are already slipping.
  - Debt capacity is small, so the deal would really be an equity-funded cash harvest.
- **What the model says a sponsor can afford.**
  - At an offer of A$10.75 (+30%), the base-case IRR is about 10%. Weighted across all five scenarios, it is about 8%.
  - Earning 20% on those weighted cash flows needs a price of about A$8.90, only about 8% above the current price.
  - Without the cost-out, the base case loses money.
- **Deal-ability is weak today.**
  - A founder-era CEO with about 6.8% of the shares, contracted to 2029.
  - A board committed to a US redomicile and to vitiligo trials.
  - A retail-heavy register anchored to historic highs.
- **Most likely outcomes.**
  - A rare-disease strategic or PE-owned platform buys it (Recordati/CVC ranks first).
  - Activists push for capital return and board renewal.
  - An opportunistic take-private becomes possible only if the oral competitors disappoint, bridged with contingent value rights (CVRs).

## Contents

| Path | What it is |
|---|---|
| [`Clinuvel_PE_Deep_Dive.md`](Clinuvel_PE_Deep_Dive.md) | Full memo: business, financials, exclusivity and competition, pipeline, governance, valuation, LBO results, structuring, buyer universe, deal mechanics, 100-day plan, diligence list, sources |
| [`model/Project_Daylight_LBO_v1.xlsx`](model/Project_Daylight_LBO_v1.xlsx) | IC-style LBO workbook: tabs MCASE / OPCASE / FINCASE / LBO / ATP / Contrib / Sensitivity / CVR; 24 cases; live formulas |
| [`model/build_model.py`](model/build_model.py) | Generates the workbook with openpyxl. Every input and its source lives at the top of the file |
| [`model/run_cases.py`](model/run_cases.py) | Runs every case headlessly in LibreOffice until each converges, filling the case-capture tables |
| [`model/make_charts.py`](model/make_charts.py) | Renders the report charts from the workbook |
| [`charts/`](charts) | Charts used in the memo |

## Rebuild

```bash
pip install openpyxl matplotlib          # LibreOffice Calc + python3-uno are also required
python3 model/build_model.py             # writes model/Project_Daylight_LBO_v1.xlsx
python3 model/run_cases.py model/Project_Daylight_LBO_v1.xlsx   # runs cases 1-24, saves cached values
python3 model/make_charts.py             # writes charts/*.png
```

In Excel, keep iterative calculation on; the file stores this setting. To change the running scenario, change `MCASE` (MCASE!C5).

## Caveats

- **Sources.** Everything comes from public sources. During the research session, direct document fetches (SEC EDGAR, clinuvel.com, ASX) were blocked, so facts were taken from search-engine extracts of the cited pages.
- **Estimates, not disclosures.** Several model inputs are analyst estimates:
  - the FY26 cost split by function;
  - the US/EU revenue split;
  - working capital and dilutive securities;
  - financing terms;
  - all forward scenarios.
- **Verify first.** Appendix B of the memo lists what to check against the FY2026 annual report and the Form 20-F before relying on the numbers.
- **Not investment advice.**
