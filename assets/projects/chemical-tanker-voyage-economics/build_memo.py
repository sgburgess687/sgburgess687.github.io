"""build_memo.py — one-page executive memo (memo.pdf) from results.json."""
import json
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph,
                                Spacer, Table, TableStyle)

HERE = Path(__file__).resolve().parent
R = json.loads((HERE / "results.json").read_text())
A = R["assumptions"]
P1 = R["route_pnl"]["1"]
TOR = R["tornado"]["1"]
sw = {r["lever"]: r["swing"] for r in TOR["rows"]}
ela = {e["lever"]: e["d_tce"] for e in R["elasticity"]["1"]}
_bi = R["breakeven"]["brent"].index(80)
BE_N = R["breakeven"]["port"]["1.00x port time"][_bi]
BE_S = R["breakeven"]["port"]["2.00x port time"][_bi]
_sp = R["speed"]
_opts = [ _sp["speeds"][max(range(len(v)), key=lambda i: v[i])] for v in _sp["curves"].values() ]
SPEED_SPREAD = max(_opts) - min(_opts)

NAVY = colors.HexColor("#1F3A5F")
INK = colors.HexColor("#1a1a1a")
INK2 = colors.HexColor("#4d4d4d")
RULE = colors.HexColor("#c9c9c4")

body = ParagraphStyle("b", fontName="Helvetica", fontSize=8.7, leading=11.6, textColor=INK,
                      alignment=TA_JUSTIFY, spaceAfter=5)
h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=9.4, leading=12, textColor=NAVY,
                   spaceBefore=7, spaceAfter=3)
small = ParagraphStyle("s", fontName="Helvetica", fontSize=7.1, leading=9, textColor=INK2)
tiny = ParagraphStyle("t", fontName="Helvetica-Oblique", fontSize=6.6, leading=8.4, textColor=INK2)
title = ParagraphStyle("T", fontName="Helvetica-Bold", fontSize=15, leading=17, textColor=NAVY)
sub = ParagraphStyle("S", fontName="Helvetica", fontSize=8.6, leading=11, textColor=INK2)


def m(x, d=0):
    return f"${x:,.{d}f}"


story = []
story.append(Paragraph("Chemical Tanker Voyage Economics — Where the Money Actually Is", title))
story.append(Paragraph(
    "Round-voyage TCE model, 38,000 dwt IMO II parcel tanker &nbsp;|&nbsp; Sean Burgess &nbsp;|&nbsp; "
    "September 2026", sub))
story.append(Spacer(1, 3))
story.append(Table([[""]], colWidths=[7.0 * inch], rowHeights=[1.6],
                   style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY)])))
story.append(Spacer(1, 7))

story.append(Paragraph("Purpose", h))
story.append(Paragraph(
    "To establish, for a mid-size deep-sea chemical parcel tanker, which variables actually determine "
    "whether a voyage earns its cost of capital — and which of those a commercial team can influence. The "
    "model prices a full round voyage (fronthaul laden, backhaul part-laden), builds up every voyage cost "
    "line, and converts the result to a time-charter-equivalent day rate.", body))

story.append(Paragraph("Base case — US Gulf to ARA round voyage", h))
kpi = [["TCE", "Voyage days", "Gross freight", "Voyage costs", "Result vs. benchmark", "Break-even rate"],
       [m(P1["tce"]) + "/day", f"{P1['days']:.1f}", f"${P1['gross']/1e6:.2f}m", f"${P1['voy_costs']/1e6:.2f}m",
        m(P1["result"]), f"${P1['be_rate_fh']:.2f}/t"]]
t = Table(kpi, colWidths=[1.16 * inch] * 6)
t.setStyle(TableStyle([
    ("FONT", (0, 0), (-1, 0), "Helvetica", 6.6), ("TEXTCOLOR", (0, 0), (-1, 0), INK2),
    ("FONT", (0, 1), (-1, 1), "Helvetica-Bold", 11), ("TEXTCOLOR", (0, 1), (-1, 1), NAVY),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("TOPPADDING", (0, 0), (-1, -1), 2),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("BOX", (0, 0), (-1, -1), 0.5, RULE),
    ("INNERGRID", (0, 0), (-1, -1), 0.4, RULE)]))
story.append(t)
story.append(Spacer(1, 3))
story.append(Paragraph(
    f"Brent at the trailing-12-month average of ${A['brent_usd_bbl']:.2f}/bbl, 85% fronthaul and 50% "
    f"backhaul utilization, normal port time, against a daily vessel-cost benchmark of "
    f"{m(A['daily_cost_benchmark'])}/day (operating expense plus capital recovery).", small))

story.append(Paragraph("Findings", h))
findings = [
    ("Commercial variables outrank fuel.", (
        f"Over plausible ranges, the freight rate swings TCE by {m(sw['Freight rate (market)'])}/day and "
        f"backhaul utilization by {m(sw['Backhaul utilization'])}/day, against {m(sw['Bunker price (Brent)'])}/day "
        f"for a P10-to-P90 move in Brent. Bunkers are the biggest cost line "
        f"({P1['bunkers']/P1['voy_costs']:.0%} of voyage costs) but rank fifth as a driver of the result.")),
    ("Cargo intake is the controllable lever.", (
        f"A 10% gain in fronthaul utilization is worth {m(ela['Fronthaul utilization'])}/day — about "
        f"{abs(ela['Fronthaul utilization']/ela['Bunker price']):.1f} times the effect of a 10% move in bunker "
        "price, and unlike the oil price it is a decision, not a market outcome. On a parcel tanker, the "
        "marginal part cargo is where the margin lives.")),
    ("Port time is a rate problem.", (
        f"The voyage spends {P1['port_days']/P1['days']:.0%} of its days in port. Severe congestion (2x port "
        f"days) costs {m(sw['Port time (congestion)'])}/day, and lifts the break-even fronthaul rate from "
        f"${BE_N:.2f}/t to ${BE_S:.2f}/t at $80 Brent — a {(BE_S/BE_N-1)*100:.0f}% uplift the market does not "
        "automatically pay. Berth productivity and laytime terms are commercial terms, not operational details.")),
    ("Route economics turn on one insurance line.", (
        f"On the ARA-India trade, Suez earns {m(R['route_pnl']['3']['tce'])}/day against "
        f"{m(R['route_pnl']['4']['tce'])}/day around the Cape. Remove the illustrative war-risk premium "
        f"(0.25% of hull value per transit) and Suez gains {m(R['route_pnl_no_war_risk']['3']['tce'] - R['route_pnl']['3']['tce'])}/day. "
        "Distance does not decide the routing; the premium does.")),
]
for lead, txt in findings:
    story.append(Paragraph(f"<b>{lead}</b> {txt}", body))


story.append(Paragraph("Sensitivity ranking — TCE swing over the plausible range of each input", h))
fmt = {"freight_mult": lambda v: f"{v:.2f}x", "brent_usd_bbl": lambda v: f"${v:.0f}",
       "port_time_mult": lambda v: f"{v:.2f}x", "util_fh": lambda v: f"{v:.0%}",
       "util_bh": lambda v: f"{v:.0%}", "speed_kn": lambda v: f"{v:.1f} kn",
       "port_cost_mult": lambda v: f"{v:.2f}x", "eua_eur": lambda v: f"EUR {v:.0f}"}
rows = [["Input", "Range tested", "TCE at low", "TCE at high", "Swing $/day", "Range set by"]]
srcs = {"Freight rate (market)": "BLS PPI YoY P10/P90", "Bunker price (Brent)": "Brent 5-yr P10/P90",
        "Port time (congestion)": "scenario band", "Fronthaul utilization": "judgment band",
        "Backhaul utilization": "judgment band", "Service speed": "operating range",
        "Port disbursements": "judgment band", "EU carbon price (EUA)": "judgment band"}
for r in TOR["rows"]:
    f = fmt[r["key"]]
    rows.append([r["lever"], f"{f(r['low'])} to {f(r['high'])}", m(r["tce_low"]), m(r["tce_high"]),
                 m(r["swing"]), srcs[r["lever"]]])
t2 = Table(rows, colWidths=[1.5 * inch, 1.25 * inch, 0.95 * inch, 0.95 * inch, 0.9 * inch, 1.45 * inch])
t2.setStyle(TableStyle([
    ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7), ("FONT", (0, 1), (-1, -1), "Helvetica", 7.4),
    ("FONT", (4, 1), (4, -1), "Helvetica-Bold", 7.4),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("BACKGROUND", (0, 0), (-1, 0), NAVY),
    ("TEXTCOLOR", (5, 1), (5, -1), INK2), ("FONT", (5, 1), (5, -1), "Helvetica-Oblique", 6.8),
    ("ALIGN", (1, 0), (4, -1), "RIGHT"), ("LINEBELOW", (0, 1), (-1, -2), 0.3, RULE),
    ("TOPPADDING", (0, 0), (-1, -1), 2.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2),
    ("BOX", (0, 0), (-1, -1), 0.5, RULE)]))
story.append(t2)
story.append(Spacer(1, 2))
story.append(Paragraph("All other inputs held at base case in each row. Every case is a complete re-run of "
                       "the voyage, not a linear approximation.", small))

story.append(Paragraph("Recommendation", h))
story.append(Paragraph(
    "Prioritize commercial effort in the order the sensitivity implies: rate negotiation and cargo "
    "intake first, port and laytime terms second, bunker and speed optimization third. Bunker hedging "
    "protects a real but secondary exposure; it does not substitute for filling the ship. Re-optimize "
    "speed per fixture rather than by policy — the TCE-maximizing speed in this model moves by "
    f"{SPEED_SPREAD:.1f} knots across the plausible bunker range. For any trade touching the EU, carry the ETS cost "
    f"({m(P1['ets'])} on this voyage, about {m(P1['ets']/P1['days'])}/day) in the rate, not in the "
    "margin.", body))

story.append(Paragraph("Assumptions and limitations — read before using any number above", h))
story.append(Paragraph(
    "<b>Simulated and estimated inputs.</b> Freight rates, port costs, port days, vessel particulars and the "
    "crude-to-bunker conversion are analyst assumptions, not observed market data. Canal tolls are built up "
    "from published Panama Canal Authority and Suez Canal Authority schedules, and US Gulf and ARA port dues "
    "from Port Houston Tariff No. 14 and the Port of Rotterdam seaport-dues schedule; Brent and the freight "
    "index are public series. Port ancillaries (pilotage, tugs, agency) and the Far East and India calls "
    "remain estimates. No public chemical-tanker rate assessment exists (Baltic and Clarksons are "
    "subscription-only), so base rates were calibrated to a plausible TCE and the BLS deep-sea freight PPI "
    "was used only to size the plausible range. Vessel-level congestion data (USACE LPMS) is not publicly "
    "available, so port time is modelled as scenario bands. Every assumption is a labelled, editable input "
    "cell in the accompanying workbook; the conclusions concern the structure of voyage economics and are "
    "robust to the level of these inputs, while the dollar figures are illustrative.", small))
story.append(Spacer(1, 5))
story.append(Paragraph(
    "Independent analysis using publicly available data. Not affiliated with or endorsed by any organization "
    "named. Model: chemical_tanker_voyage_model.xlsx (five scenario switches, full cost build-up, tornado, "
    "break-even curves); code: analysis.py; sources: data/README.md.", tiny))

doc = BaseDocTemplate(str(HERE / "memo.pdf"), pagesize=LETTER,
                      leftMargin=0.75 * inch, rightMargin=0.75 * inch,
                      topMargin=0.6 * inch, bottomMargin=0.55 * inch, title="Chemical Tanker Voyage Economics — memo",
                      author="Sean Burgess")
doc.addPageTemplates([PageTemplate(id="p", frames=[Frame(0.75 * inch, 0.55 * inch, 7.0 * inch, 10.0 * inch,
                                                         leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)])])
doc.build(story)
print("pages check -> memo.pdf written")
