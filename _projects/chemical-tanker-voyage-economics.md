---
title: "Chemical Tanker Voyage Economics"
domain: "Maritime Commercial Analytics"
date: 2026-09-17
summary: >-
  A round-voyage TCE model for a 38,000 dwt chemical parcel tanker, built to answer which
  variables actually decide whether a voyage earns its cost of capital — and which of them a
  commercial team can control.
role: "Independent analysis"
timeframe: "September 2026"
tools: ["Excel", "Python", "pandas", "openpyxl", "Chart.js"]
data_sources: "FRED/EIA Brent, BLS deep-sea freight PPI, Panama and Suez canal tariffs, Port Houston Tariff No. 14, Port of Rotterdam seaport dues"
image: /assets/projects/chemical-tanker-voyage-economics/preview.png
demo: /assets/projects/chemical-tanker-voyage-economics/index.html
github: https://github.com/sgburgess687/sgburgess687.github.io/tree/master/assets/projects/chemical-tanker-voyage-economics
description: "Round-voyage time-charter-equivalent model for a chemical parcel tanker: full voyage cost build-up from published canal and port tariffs, tornado sensitivity, and break-even freight curves."
---

*Independent analysis using publicly available data. Not affiliated with or endorsed by any
organization named. Freight rates, port ancillaries, port days and vessel particulars are analyst
assumptions and are labelled as such throughout.*

## The question

A voyage charter pays a freight rate in dollars per tonne. A shipowner's alternative — putting the
same ship on time charter — pays in dollars per day. The bridge between them is the time charter
equivalent: gross freight, less commissions, less every cost the owner bears on the voyage, divided
by the days the voyage consumes. Every chartering desk runs that calculation before fixing a cargo.

The interesting question is not what the number is. It is what moves it. Fuel dominates the
conversation in shipping because it is the biggest cost line — so I wanted to test whether it is
also the biggest *driver*, and to separate the variables a commercial team can influence from the
ones it can only absorb.

## What I built

A round-voyage model for a generic 38,000 dwt IMO II stainless parcel tanker across four trades —
US Gulf–ARA, US Gulf–Far East via Panama, and ARA–India West Coast via Suez and around the Cape —
with a fronthaul cargo and a backhaul cargo at lower utilization.

The cost build-up is assembled from published schedules wherever one exists:

- **Bunkers** split between VLSFO at sea and MGO inside emission control areas and in port, with
  consumption scaling on the cube of speed.
- **Canal tolls** built from the Panama Canal Authority's chemical-tanker capacity tariff and the
  Suez Canal Authority's SDR-per-net-ton bands, including the laden/ballast split and the
  chemical-tanker surcharge.
- **Port costs** from the actual tariffs: Port Houston's $813.40 harbour fee and its dockage rate
  per foot of LOA per day, on the published taper that steps down after the second day; Rotterdam's
  seaport dues as a GT component plus a cargo component capped at 133.3% of GT, plus the
  sustainability component.
- **EU ETS** allowances at the 2026 full phase-in, covering half of the emissions on a voyage into
  or out of Europe plus all of the emissions at an EU berth.

The deliverable is an Excel workbook with five scenario switches and a row-wise calculation engine
behind every chart, so each sensitivity case is a complete re-run of the voyage rather than a linear
approximation. The same engine is implemented independently in Python, and the workbook is checked
against it cell for cell.

![Tornado chart ranking the drivers of TCE](/assets/projects/chemical-tanker-voyage-economics/preview.png)

## What it found

**Commercial variables outrank fuel.** Across plausible ranges — the freight band from the
distribution of the BLS deep-sea freight index, the bunker band from Brent's five-year P10 to P90 —
the freight rate swings TCE by $15,266 a day and backhaul utilization by $13,259, against $5,934 for
the oil price. Bunkers are 62% of voyage costs and the fifth-ranked driver of the result.

**Cargo intake is the lever an operator owns.** A 10% gain in fronthaul utilization is worth $3,130
a day, roughly 2.4 times a 10% move in bunker price — and unlike the oil price, it is a decision.
On a parcel tanker, that is the entire commercial game: the marginal part cargo that fills the last
few tanks.

**Port time is a rate problem, not an operations footnote.** The base voyage spends 34% of its days
in port. Severe congestion doubles port time, costs $9,200 a day of TCE, and lifts the break-even
freight rate from $42.23 to $56.35 per tonne — a 33% uplift the market will not automatically pay.

**Route economics turn on one insurance line.** On the ARA–India trade, Suez earns $21,396 a day
against $16,496 around the Cape. Remove the war-risk premium and Suez gains another $3,009. Distance
does not decide that routing; the premium does.

## How to judge it

The honest limitations are in the write-up and the memo rather than buried. There is no public
chemical-tanker freight assessment — the Baltic and Clarksons series are subscription-only — so base
rates are calibrated to a plausible TCE and the public freight index is used only to size the
*range*, not the level. Port ancillaries (pilotage, tugs, agency) and the Far East and India calls
are estimates, as are port days; vessel-level berth and lock delay data is not publicly available,
so congestion is modelled as scenario bands. Bunker prices are derived from crude rather than
observed, which understates the risk of a widening distillate crack.

None of that undermines the structural finding, which is what the model is for: the ranking of the
drivers is robust to the level of the inputs, and every input is a labelled cell a reader can
disagree with and change.

## The artifacts

- **[Full interactive write-up](/assets/projects/chemical-tanker-voyage-economics/index.html)** — the analysis with live charts
- **[Excel model](/assets/projects/chemical-tanker-voyage-economics/model/chemical_tanker_voyage_model.xlsx)** — five scenario switches, full cost build-up, tornado, break-even curves
- **[One-page memo (PDF)](/assets/projects/chemical-tanker-voyage-economics/memo.pdf)** — findings and recommendation
- **[analysis.py](/assets/projects/chemical-tanker-voyage-economics/analysis.py)** and **[data sources](/assets/projects/chemical-tanker-voyage-economics/data/README.md)** — the code and every source with its retrieval date
