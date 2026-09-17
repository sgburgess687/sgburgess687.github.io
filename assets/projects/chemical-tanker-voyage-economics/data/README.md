# Data sources — Chemical Tanker Voyage Economics Model

*Independent analysis using publicly available data. Not affiliated with or endorsed by any organization named.*

All files in this folder are public data. Published tariff schedules also feed the cost build-up directly:
Panama and Suez canal tolls, Port Houston's harbor fee and dockage, and Port of Rotterdam seaport dues
(section 2 below). **Everything else — freight rates, port ancillaries and the Far East/India port calls,
port days, vessel particulars, the crude-to-bunker conversion, canal net tonnages, the war-risk premium and
the daily vessel-cost benchmark — is an analyst assumption, not observed data.** Those assumptions are
listed on the `Inputs`, `Routes`, `Port_Tariffs` and `Canal_Tolls` sheets of the workbook, each with its
basis written beside it, and in the "Assumptions" table of `index.html`.

---

## 1. Files in this folder

### `fred_brent_monthly_MCOILBRENTEU.csv`
- **What:** Europe Brent crude spot price FOB, monthly average, US$/bbl, Jan-2015 to Aug-2026 (140 observations).
- **Source:** U.S. Energy Information Administration, republished by FRED (Federal Reserve Bank of St. Louis),
  series `MCOILBRENTEU`.
- **URL:** https://fred.stlouisfed.org/series/MCOILBRENTEU (text file: https://fred.stlouisfed.org/data/MCOILBRENTEU.txt)
- **Retrieved:** 17 September 2026. FRED reported "Last Updated: 2026-09-02".
- **Used for:** the bunker-price driver (base case = trailing-12-month average; tornado range = 5-year P10/P90).
- **Verified:** this file was checked against the CSV downloaded directly from FRED on 17 September 2026
  (`sources/fred_brent_monthly.csv` in the project root). All 140 monthly observations match to the cent;
  zero differences. Base-case values: latest month (Aug-2026) $91.08; trailing-12-month average $82.01;
  5-year average $83.88; 5-year P10 $67.98; 5-year P90 $104.84.

### `bls_ppi_deep_sea_freight_WPU30130101.csv`
- **What:** U.S. Producer Price Index, Deep Sea Freight Transportation, monthly index (Dec-2008 = 100),
  Dec-2008 to Jun-2026 (211 observations).
- **Source:** U.S. Bureau of Labor Statistics series `WPU30130101`, via FRED.
- **URL:** https://fred.stlouisfed.org/series/WPU30130101 · BLS: https://data.bls.gov/timeseries/WPU30130101
- **Retrieved:** September 2026 (file supplied with the project data package, `tanker_voyage_data/freight_rates/`).
- **Used for:** sizing the plausible freight-rate range only — the P10/P90 of year-on-year change since 2016
  (−8.9% / +24.9%) become the low and high freight multipliers in the tornado.
- **Not used for:** the level of freight rates. This is a broad U.S. deep-sea freight index covering all
  cargo types; it is not a chemical-tanker rate. Spot chemical rates are considerably more volatile, so the
  freight band in this model is probably conservative.

---

## 2. Tariff sources used in the cost build-up (values transcribed into the workbook)

### Panama Canal tolls
- **Source:** Panama Canal Authority (ACP), consolidated maritime services tariff schedule
  (`PP-Lista-tarifas-consolidado-final-ingles.pdf`).
- **URL:** https://pancanal.com/en/tolls · PDF: https://pancanal.com/wp-content/uploads/2021/08/PP-Lista-tarifas-consolidado-final-ingles.pdf
- **Retrieved:** September 2026 (extract supplied in `tanker_voyage_data/canal_tolls/`).
- **Items used:** fixed tariff per transit, "Super" vessel, all cargo types (1010.FS01, $100,000); capacity
  tariff per PC/UMS ton for chemical tankers, Super vessel (1010.QS01, $4.00); ballast discount (1010.BA01,
  85% of tolls); mandatory tug; linehandling; transit reservation; security charge; fresh-water surcharge
  (fixed portion); port pilotage.
- **Judgment applied:** the vessel is classified as a "Super" vessel (beam over 91 ft) for a 38,000 dwt
  chemical tanker; PC/UMS tonnage of 20,000 is an estimate; the variable portion of the fresh-water surcharge
  and the Neopanamax-locks items are omitted. ACP restructured its tariff in 2024 and further changes have
  been reported in 2026 trade press — **re-download the PDF before using these figures for anything real.**

### Suez Canal tolls
- **Source:** Suez Canal Authority Circular 7/2023, effective 15 January 2024 (petroleum-product-tanker
  SDR/SCNT bands, laden and ballast).
- **URL:** official tolls page https://www.suezcanal.gov.eg/English/Navigation/Tolls/Pages/TollsTable.aspx ·
  circular as mirrored by a shipping agency:
  https://www.kanooshipping.com/storage/uploads/circulars/1697703461_SUEZ%20CANAL%20-%20REVISED%20TOLL%20FEE%20-%20Circular%207%202023%20-%201.pdf
- **Retrieved:** September 2026 (extract supplied in `tanker_voyage_data/canal_tolls/`).
- **Judgment applied:** the SCA does not publish a chemical-tanker band in this circular, so the
  petroleum-product-tanker bands are used with the 20% chemical-tanker surcharge reported for the same
  circular (15% for ballast). SCNT of 20,000 is an estimate; the SDR/USD rate (1.35) is an estimate — the
  IMF publishes the daily rate at https://www.imf.org/external/np/fin/data/rms_five.aspx. Ancillary transit
  costs ($40,000) and the war-risk additional premium (0.25% of hull value) are estimates and are labelled
  as such in the workbook. **Suez rates have changed repeatedly since this circular — confirm the current
  one before use.**

### Port Houston — Tariff No. 14 (effective 1 January 2025)
- **What:** harbor fee and dockage rates, used to build the US Gulf port cost.
- **Source:** Port of Houston Authority Tariff No. 14, Subrule 096 (Harbor Fee) and Subrule 095 (Dockage).
- **URL:** https://porthouston.com/toolbox/rates/tariffs/ · PDF: https://porthouston.com/wp-content/uploads/2024/12/v2-Tariff-14-January-1-2025-12.5.24.pdf
- **Retrieved:** 17 September 2026 (PDF in `sources/port_houston_tariff14_2025.pdf`).
- **Rates used:** harbor fee **$813.40** per vessel entry for vessels 250 ft and over; dockage **$11.24 per
  foot of LOA per 24 hours** in the 600–650 ft band, with the published taper — days 1 and 2 at the full
  rate, day 3 at 90%, day 4 at 75%, day 5 at 60%, day 6 onward at 50%.
- **Judgment applied:** the modelled vessel is 600 ft LOA; 60% of US Gulf port days are assumed to be
  alongside (dockage accrues) and the rest waiting at anchorage (no dockage) — that split is an estimate.
  Port Houston dockage applies at Port Authority wharves; chemical parcels frequently load at private
  liquid terminals whose dockage differs. Pilotage, tugs, linesmen, agency and terminal charges are an
  estimated $45,000 per call.

### Port of Rotterdam — seaport dues
- **What:** the GT, cargo and sustainability components of Rotterdam seaport dues, used for the ARA port cost.
- **Source:** Port of Rotterdam, "General terms and conditions including port tariffs", Annex 1 §1.2,
  Tables 1–3, and the efficiency-discount cap in §1.4(A).
- **URL:** https://www.portofrotterdam.com/en/port-dues-tariffs · PDF (2025 schedule):
  https://www.portofrotterdam.com/sites/default/files/2026-01/general-terms-and-conditions-including-port-tariffs-2025.pdf
- **Retrieved:** 17 September 2026 (PDF in `sources/rotterdam_tariffs_2025.pdf`).
- **Rates used:** chemical/gas tanker GT tariff **€0.303 per GT** (Table 1, type C); "other liquid bulk"
  cargo rate **€0.576 per tonne** (Table 2, commodity 07); sustainability component **€0.065 per GT**
  (Table 3); cargo component capped at **GT × 133.3% × cargo rate** (§1.4(A), chemical/gas tankers).
- **Judgment applied:** the 2025 schedule is indexed by the **3.5%** 2026 indexation the port published, as
  the 2026 PDF was not retrievable at the same URL pattern. No ESI or Green Award discount is taken — a
  Green Award certificate would cut the sustainability component by 70%. GT of 23,500 is an estimate for
  the class (the port's own worked example uses a 23,230 GT chemical tanker). Pilotage, tugs, boatmen,
  agency and terminal charges are an estimated $35,000 per call.

---

## 3. What could not be obtained, and what was done instead

| Needed | Status | Substitute used |
|---|---|---|
| Chemical-tanker spot freight rates ($/t) | Not public. Baltic Exchange and Clarksons assessments are subscription-only. | Base rates calibrated to a plausible TCE for the vessel class; BLS PPI used only to size the range. **This is the largest assumption in the model.** |
| VLSFO / MGO bunker prices by port | Not free. Ship & Bunker CSV export, Platts and Argus are paid; the Baltic bunker report is members-only. | Modelled as fixed multiples of Brent (VLSFO = Brent × 7.7, MGO = Brent × 9.7), both editable inputs. Spot-check against Ship & Bunker's free charts: https://shipandbunker.com |
| Port disbursement accounts (Houston, Rotterdam, Ulsan, Mumbai) | Partly solved. The Houston and Rotterdam **published tariffs** are now built into the model (see section 2). What remains confidential is the ancillary stack — pilotage, tugs, linesmen, agency, private-terminal charges — and the Far East and India calls, for which no schedule was transcribed. | Published rates where they exist; estimated ancillaries ($45k Houston, $35k Rotterdam per call) and lump sums for Far East ($65k) and India ($55k), with a ±25% band in the tornado. |
| Berth waiting time and inland lock delay | Restricted. USACE Lock Performance Monitoring System vessel-level data "is only accessible on USACE ACE-IT Machines; non-USACE users must submit a FOIA request" (https://www.cisa.gov/mts-resilience-resources/lock-performance-monitoring-system-lpms). | Port time modelled as scenario bands (normal / congested / severe), which is also how commercial voyage estimators handle it. Lock physical characteristics are open data if needed: https://geospatial-usace.opendata.arcgis.com |
| Sea distances | No free authoritative table was pulled. | Approximate standard-routeing figures (Houston–Rotterdam ~5,000 nm; Houston–Ulsan via Panama ~9,650 nm; Rotterdam–Mumbai ~6,350 nm via Suez, ~10,850 nm via Cape). Verify against a distance table (Netpas, BP Shipping Marine Distances) before relying on them. |
| EU ETS allowance price | Public but not pulled here. | €75/t CO₂ assumed; check ICE EUA futures. Emission factors (3.114 t CO₂/t VLSFO, 3.206 t CO₂/t MGO) are the standard EU MRV / IMO factors, and the 2026 phase-in is 100% of covered emissions. |

---

## 4. Reproducing the analysis

```bash
python analysis.py        # loads the two CSVs, runs the voyage engine, writes results.json
python build_workbook.py  # builds model/chemical_tanker_voyage_model.xlsx (live formulas)
python build_site.py      # renders index.html from results.json
python build_memo.py      # renders memo.pdf from results.json
```

The workbook and `analysis.py` implement the same engine independently; the workbook's `Voyage_PnL` sheet
carries a check cell that compares its TCE with the calculation engine's, and the two were verified to agree
to six decimal places across several scenario combinations.

## 5. Raw source files

The PDFs and CSV downloaded for this project are kept outside the published folder, in `sources/` at the
project-data root: `fred_brent_monthly.csv`, `port_houston_tariff14_2025.pdf`, `rotterdam_tariffs_2025.pdf`.

*Retrieved and compiled 17 September 2026.*
