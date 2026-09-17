"""build_site.py — renders index.html from results.json (run analysis.py first)."""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
R = json.loads((HERE / "results.json").read_text())
A, RG = R["assumptions"], R["ranges"]
P1 = R["route_pnl"]["1"]
TOR = R["tornado"]["1"]
ELA = R["elasticity"]["1"]

PAY = json.dumps({
    "routes": {k: {"name": v["route"], "tce": v["tce"], "days": v["days"], "gross": v["gross"],
                   "bunkers": v["bunkers"], "port": v["port"], "canal": v["canal"], "ets": v["ets"],
                   "misc": v["misc"], "result": v["result"], "be": v["be_rate_fh"]}
               for k, v in R["route_pnl"].items()},
    "tornado": TOR, "elasticity": ELA, "breakeven": R["breakeven"], "grid": R["grid"],
    "speed": R["speed"], "brent": R["brent_series"], "ppi": R["ppi_series"], "p1": P1,
    "benchmark": A["daily_cost_benchmark"],
}, separators=(",", ":"))


def money(x, d=0):
    return f"${x:,.{d}f}"


HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Chemical Tanker Voyage Economics</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{
  --surface-0:#ffffff; --surface-1:#fcfcfb; --surface-2:#f4f4f2; --line:#e3e3df;
  --ink:#0b0b0b; --ink-2:#52514e; --ink-3:#7a7975; --navy:#1f3a5f;
  --s1:#2a78d6; --s2:#eb6834; --s3:#1baf7a; --s4:#eda100; --neg:#e34948;
}
*{box-sizing:border-box}
body{margin:0;background:var(--surface-1);color:var(--ink);
  font:16px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;}
.wrap{max-width:980px;margin:0 auto;padding:0 20px 80px}
header{background:var(--navy);color:#fff;padding:44px 0 36px;margin-bottom:28px}
header .wrap{padding-bottom:0}
h1{font-size:30px;line-height:1.2;margin:0 0 6px;font-weight:650}
header p.sub{margin:0;color:#c9d6e6;font-size:16px}
header p.by{margin:14px 0 0;color:#9fb4cd;font-size:13px}
h2{font-size:21px;margin:44px 0 6px;color:var(--navy);font-weight:650}
h3{font-size:15px;margin:26px 0 6px;color:var(--ink);font-weight:650}
p,li{color:var(--ink-2)}
.lede{font-size:17px;color:var(--ink)}
.disclaimer{background:#fdf3f1;border:1px solid #f3cdc4;border-left:4px solid var(--s2);
  padding:12px 16px;border-radius:6px;font-size:13.5px;color:#6b3a2c;margin:18px 0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:22px 0}
.kpi{background:var(--surface-0);border:1px solid var(--line);border-radius:8px;padding:14px 16px}
.kpi .v{font-size:24px;font-weight:650;color:var(--navy);letter-spacing:-.4px}
.kpi .k{font-size:12px;color:var(--ink-3);text-transform:uppercase;letter-spacing:.04em;margin-top:4px}
.card{background:var(--surface-0);border:1px solid var(--line);border-radius:8px;padding:18px 18px 10px;margin:18px 0}
.card h4{margin:0 0 2px;font-size:15px;font-weight:650}
.card .note{margin:0 0 12px;font-size:13px;color:var(--ink-3)}
.chart{position:relative;height:320px}
.chart.tall{height:400px}
table{border-collapse:collapse;width:100%;font-size:13.5px;margin:10px 0 4px}
th,td{padding:7px 9px;border-bottom:1px solid var(--line);text-align:right}
th:first-child,td:first-child{text-align:left}
thead th{background:var(--surface-2);color:var(--ink);font-weight:650;font-size:12.5px;
  text-transform:uppercase;letter-spacing:.03em;border-bottom:1px solid var(--line)}
tbody tr:hover{background:var(--surface-1)}
td.num{font-variant-numeric:tabular-nums}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:760px){.grid2{grid-template-columns:1fr}}
.tag{display:inline-block;font-size:11px;letter-spacing:.04em;text-transform:uppercase;
  background:var(--surface-2);color:var(--ink-3);border-radius:4px;padding:2px 7px;margin-left:6px}
.assume{font-size:13.5px}
.assume td:last-child{color:var(--ink-3)}
.assume td:not(:first-child){text-align:left}
footer{border-top:1px solid var(--line);margin-top:56px;padding-top:18px;font-size:13px;color:var(--ink-3)}
a{color:var(--s1)}
ul{padding-left:20px}
.hm td{text-align:center;font-variant-numeric:tabular-nums;border:2px solid var(--surface-0)}
.dl a{display:inline-block;margin:6px 10px 6px 0;padding:9px 14px;border:1px solid var(--line);
  border-radius:6px;background:var(--surface-0);text-decoration:none;color:var(--navy);font-size:14px;font-weight:600}
</style>
</head>
<body>
<header><div class="wrap">
  <h1>Chemical Tanker Voyage Economics</h1>
  <p class="sub">A round-voyage TCE model for a 38,000 dwt IMO II parcel tanker — cost build-up, sensitivity, and break-even freight</p>
  <p class="by">Sean Burgess · September 2026 · Houston, TX</p>
</div></header>
<div class="wrap">

<p class="lede">This is a working voyage-estimation model, the kind a chartering desk runs before fixing a
cargo. It converts a voyage charter into a time-charter-equivalent (TCE) day rate, breaks the cost stack
apart, and then asks the question that matters commercially: <strong>of everything that moves a voyage
result, which levers actually matter, and which of those can an operator control?</strong></p>

<div class="disclaimer"><strong>Independent analysis using publicly available data. Not affiliated with or
endorsed by any organization named.</strong><br>
Canal tolls are taken from published tariff schedules and the Brent and freight-index series are public
data. <strong>Freight rates, port costs, port days, vessel particulars and the crude-to-bunker conversion
are analyst assumptions, not observed market data</strong> — every one of them is stated explicitly below and
is an editable input in the Excel model. The conclusions are about the <em>structure</em> of voyage
economics, which is robust to the exact level of these inputs; the dollar figures are illustrative.</div>

<h2>Headline result — base case</h2>
<p>Base case: US Gulf → ARA round voyage, 85% fronthaul / 50% backhaul utilization, Brent at the
trailing-12-month average of $__BRENT__/bbl, normal port time.</p>
<div class="kpis">
  <div class="kpi"><div class="v">__TCE__</div><div class="k">TCE per day</div></div>
  <div class="kpi"><div class="v">__DAYS__</div><div class="k">Round-voyage days</div></div>
  <div class="kpi"><div class="v">__GROSS__</div><div class="k">Gross freight</div></div>
  <div class="kpi"><div class="v">__COSTS__</div><div class="k">Voyage costs</div></div>
  <div class="kpi"><div class="v">__BE__</div><div class="k">Break-even rate ($/t)</div></div>
</div>

<h3>What the model says</h3>
<ul>
<li><strong>Commercial decisions dominate, not fuel.</strong> Across plausible ranges, freight rate and
cargo utilization swing TCE by __SW1__ and __SW2__ per day respectively, while a P10-to-P90 move in Brent
($__BP10__ to $__BP90__) swings it by __SWB__. Bunkers are the largest single cost line
(__BUNKSHARE__% of voyage costs) but only the fifth-ranked driver of the result, because freight rates and
cargo intake move by more, in percentage terms, than the oil price does.</li>
<li><strong>Filling the ship is the lever an operator actually owns.</strong> A 10% improvement in
fronthaul utilization is worth __ELA_UTIL__/day — roughly __UTIL_VS_BUNKER__ the effect of a 10% move in
bunker price, and unlike the oil price it is inside the commercial team's control. On a parcel tanker this
is the whole commercial game: finding the part cargo that fills the last four tanks.</li>
<li><strong>Port time is the hidden cost line.</strong> This voyage spends __PORTSHARE__% of its days in
port, not at sea. Moving from normal to severe congestion (2× port days) costs __PORTSWING__/day of TCE —
more than the entire plausible range of bunker prices. Almost all of that is the denominator: the same
freight earned over more days, plus some extra MGO burn at berth. Time is the cost.</li>
<li><strong>Speed is a trade-off that moves with the oil price.</strong> In this model the
TCE-maximizing speed falls from __OPTSPEED_LOW__ knots at $65 Brent to __OPTSPEED__ knots at the base price
and __OPTSPEED_HIGH__ knots at $104 — bunkers rise with the cube of speed while days fall only in
proportion. Two caveats: the cube law flatters deep slow steaming, and engine-load limits, charterparty
speed warranties and laycan windows bind in practice. The finding that survives both is directional — the
optimum is a function of the bunker price, so it belongs in the fixture calculation rather than in a
standing speed policy.</li>
</ul>

<h2>The voyage and the cost stack</h2>
<div class="grid2">
  <div class="card"><h4>Voyage cost build-up by route</h4>
    <p class="note">Same vessel, same levers, four trade routes.</p>
    <div class="chart"><canvas id="costs"></canvas></div></div>
  <div class="card"><h4>TCE vs. daily vessel cost benchmark</h4>
    <p class="note">Benchmark __BENCH__/day (OPEX plus capital recovery, or T/C-in hire).</p>
    <div class="chart"><canvas id="tce"></canvas></div></div>
</div>

<table>
<thead><tr><th>Route</th><th>Days</th><th>Gross freight</th><th>Bunkers</th><th>Port</th><th>Canal</th>
<th>EU ETS</th><th>TCE $/day</th><th>Break-even $/t</th></tr></thead>
<tbody>__ROUTETABLE__</tbody>
</table>
<p style="font-size:13px;color:var(--ink-3)">Port cost on the two Atlantic legs is built from the published
tariffs rather than a lump sum: Port Houston charges a fixed harbour fee plus dockage per foot of LOA per
day on a taper, so congestion raises the cash cost as well as the time cost, while Rotterdam charges per GT
plus per tonne of cargo — capped — so a fuller ship pays more in dues but far more in freight. The Suez route carries a war-risk additional premium set at
0.25% of hull value per transit, an illustrative figure; set it to zero in the model and the Suez routing
gains __WARRISK__/day of TCE, widening its advantage over the Cape alternative to __SUEZGAP__/day. That single input — not distance — is what has
decided Suez-versus-Cape routing decisions since 2023.</p>

<h2>What drives the result</h2>
<div class="card"><h4>Tornado — change in TCE across a plausible range for each input</h4>
  <p class="note">Freight and bunker ranges are data-derived (BLS PPI year-on-year P10/P90; Brent 5-year
  P10/P90). Port time, utilization, speed and port-cost ranges are judgment bands, stated in the model.</p>
  <div class="chart tall"><canvas id="tornado"></canvas></div></div>

<table><thead><tr><th>Input</th><th>Low</th><th>High</th><th>TCE at low</th><th>TCE at high</th>
<th>Swing $/day</th></tr></thead><tbody>__TORTABLE__</tbody></table>

<div class="card"><h4>Leverage — TCE change for a +10% move in each input</h4>
  <p class="note">The tornado answers "what moved results most over realistic ranges". This answers
  "where does one unit of effort buy the most TCE" — the ranges differ, so both views are needed.</p>
  <div class="chart"><canvas id="elas"></canvas></div></div>

<h2>Break-even freight rates</h2>
<p>The break-even fronthaul rate is the $/t at which the voyage earns exactly the daily vessel-cost
benchmark. It is linear in the bunker price, so the slope of each line is the bunker exposure of the trade,
and the gap between lines is the cost of congestion or of sailing part-empty.</p>
<div class="grid2">
  <div class="card"><h4>By port time</h4><p class="note">USG–ARA, base utilization.</p>
    <div class="chart"><canvas id="bePort"></canvas></div></div>
  <div class="card"><h4>By fronthaul utilization</h4><p class="note">USG–ARA, normal port time.</p>
    <div class="chart"><canvas id="beUtil"></canvas></div></div>
</div>
<p>Two readings matter commercially. First, congestion is a rate problem: at severe congestion the
US Gulf–ARA voyage needs __BE_SEVERE__/t instead of __BE_NORMAL__/t at the base bunker price — a
__BE_UPLIFT__% uplift the market will not automatically pay. Second, the utilization curves are further
apart than the port-time curves, which is the same message as the tornado: cargo intake, not fuel, is what
decides whether the voyage clears its cost of capital.</p>

<div class="card"><h4>TCE ($/day) by freight level and bunker price</h4>
  <p class="note">USG–ARA. Shaded against the __BENCH__/day benchmark: values at or below it are voyages
  that do not cover the vessel.</p>
  __HEATMAP__
</div>

<div class="card"><h4>TCE versus service speed at three bunker prices</h4>
  <p class="note">Consumption scales with the cube of speed; days scale inversely. The optimum moves.</p>
  <div class="chart"><canvas id="speed"></canvas></div></div>

<h2>Assumptions — state them, so they can be argued with</h2>
<p>A voyage estimate is only as good as its inputs, and the honest thing to do is put them where a
commercial analyst can disagree with them. Every value below is a labelled input cell in the workbook.</p>
<table class="assume"><thead><tr><th>Input</th><th>Value</th><th>Basis</th></tr></thead>
<tbody>__ASSUMPTIONS__</tbody></table>

<h2>Method</h2>
<p>TCE = (gross freight − commissions − voyage costs) ÷ voyage days, for a round voyage: fronthaul A→B
laden, backhaul B→A at lower utilization and a lower rate. Voyage costs are the owner's account on a
voyage charter: bunkers (VLSFO at sea outside emission control areas, MGO inside ECAs and in port), port
disbursements, canal tolls, EU ETS allowances, and miscellaneous cargo costs. Fuel consumption scales with
the cube of speed; distance carries a 5% weather margin; EU ETS covers 50% of emissions on voyages into or
out of the EU plus 100% at an EU berth, at the 2026 100% phase-in.</p>
<p>The Excel workbook is the primary artifact: five scenario switches (route, bunker case, freight market,
congestion, utilization), a full build-up on the P&amp;L sheet, and a row-wise calculation engine behind
every chart, so each sensitivity case is a complete re-run of the voyage rather than a linear
approximation. <code>analysis.py</code> implements the same engine in Python and the workbook is checked
against it cell for cell (the model's integrity check cell on the P&amp;L sheet).</p>

<h2>Limitations</h2>
<ul>
<li><strong>Freight rates are calibrated assumptions, not fixtures.</strong> Public chemical-tanker rate
assessments do not exist — the Baltic and Clarksons series are subscription-only. Base rates were set so
that the base case earns a plausible TCE for this vessel class; the BLS deep-sea freight PPI is used only
to size the plausible <em>range</em>, not the level. Spot chemical rates are more volatile than that index,
so the freight bar in the tornado is, if anything, understated.</li>
<li><strong>Port costs are half published, half estimated.</strong> The US Gulf and ARA legs are built from the
actual published schedules — Port Houston's harbour fee and LOA-based dockage taper, and Rotterdam's GT,
cargo and sustainability components with the 133.3% chemical-tanker cap. What stays an estimate is the
ancillary stack every port has and no port publishes: pilotage, tugs, linesmen, agency and terminal
charges, plus the Far East and India calls entirely. Port days are estimates throughout. A ±25% band on
port costs is in the tornado for that reason.</li>
<li><strong>Bunker prices are derived from crude, not observed.</strong> VLSFO and MGO are modelled as
fixed multiples of Brent. In reality the crack spread moves on its own, and a widening distillate crack
would hit this trade harder than the model shows, because MGO covers all ECA and port consumption.</li>
<li><strong>Congestion is a scenario band, not measured delay.</strong> Vessel-level lock and berth delay
microdata (USACE LPMS) is not publicly downloadable; it requires a FOIA or district data request. Treating
port time as a scenario is how commercial estimators handle it too, but it is an assumption, not a
measurement.</li>
<li><strong>One ship, one parcel structure.</strong> Real parcel tankers carry many grades for many
charterers with segregation constraints and overlapping port rotations. This model treats the voyage as a
fronthaul and a backhaul parcel, which is the right level for rate and routing decisions but understates
both the complexity and the revenue upside of true parcel optimization.</li>
</ul>

<h2>Data sources</h2>
<ul>
<li><strong>Brent crude spot price, monthly</strong> — FRED series <code>MCOILBRENTEU</code> (EIA), Jan-2015
to Aug-2026, retrieved 17-Sep-2026. <a href="https://fred.stlouisfed.org/series/MCOILBRENTEU">fred.stlouisfed.org/series/MCOILBRENTEU</a></li>
<li><strong>PPI, deep sea freight transportation</strong> — BLS series <code>WPU30130101</code> via FRED,
Dec-2008 to Jun-2026. <a href="https://fred.stlouisfed.org/series/WPU30130101">fred.stlouisfed.org/series/WPU30130101</a></li>
<li><strong>Panama Canal tolls</strong> — Panama Canal Authority consolidated tariff schedule (chemical
tanker line items). <a href="https://pancanal.com/en/tolls">pancanal.com/en/tolls</a></li>
<li><strong>Suez Canal tolls</strong> — SCA Circular 7/2023 (effective 15-Jan-2024), petroleum-product
tanker bands plus the chemical-tanker surcharge. <a href="https://www.suezcanal.gov.eg/English/Navigation/Tolls/Pages/TollsTable.aspx">suezcanal.gov.eg</a></li>
<li><strong>Port Houston tariff</strong> — Port of Houston Authority Tariff No. 14, effective 1 January 2025;
Subrule 095 (dockage) and Subrule 096 (harbor fee). <a href="https://porthouston.com/toolbox/rates/tariffs/">porthouston.com/toolbox/rates/tariffs</a></li>
<li><strong>Port of Rotterdam seaport dues</strong> — General terms and conditions including port tariffs,
Annex 1 Tables 1-3 and the 133.3% chemical-tanker efficiency cap; 2026 indexation 3.5%.
<a href="https://www.portofrotterdam.com/en/port-dues-tariffs">portofrotterdam.com/en/port-dues-tariffs</a></li>
<li><strong>Inland/berth delay</strong> — USACE LPMS microdata is restricted; modelled as scenario bands.
<a href="https://www.cisa.gov/mts-resilience-resources/lock-performance-monitoring-system-lpms">cisa.gov</a></li>
</ul>
<p style="font-size:13px">Full retrieval notes, caveats and the honest account of what could not be
obtained are in <code>data/README.md</code>.</p>

<h2>Files</h2>
<p class="dl">
<a href="model/chemical_tanker_voyage_model.xlsx">Excel model (.xlsx)</a>
<a href="memo.pdf">One-page memo (PDF)</a>
<a href="analysis.py">analysis.py</a>
<a href="data/README.md">Data sources</a>
</p>

<footer>Independent analysis using publicly available data. Not affiliated with or endorsed by any
organization named. Freight rates, port costs and vessel particulars are analyst assumptions.
Sean Burgess · September 2026</footer>
</div>

<script>
const D = __DATA__;
const INK='#52514e', INK3='#7a7975', LINE='#e3e3df';
Chart.defaults.font.family='-apple-system,BlinkMacSystemFont,Segoe UI,Arial,sans-serif';
Chart.defaults.font.size=12; Chart.defaults.color=INK;
Chart.defaults.plugins.legend.labels.boxWidth=10; Chart.defaults.plugins.legend.labels.boxHeight=10;
Chart.defaults.plugins.legend.labels.usePointStyle=true;
const S={b:'#2a78d6',o:'#eb6834',g:'#1baf7a',y:'#eda100',r:'#e34948',v:'#4a3aa7'};
const usd=v=>'$'+Math.round(v).toLocaleString();
const gridX={grid:{color:LINE,drawTicks:false},border:{display:false}};
const gridY={grid:{display:false},border:{display:false}};
const names=Object.values(D.routes).map(r=>r.name.replace(' via',String.fromCharCode(10)+'via').replace(' transatlantic',String.fromCharCode(10)+'transatlantic'));

new Chart(document.getElementById('costs'),{type:'bar',data:{labels:names,datasets:[
 {label:'Bunkers',data:Object.values(D.routes).map(r=>r.bunkers),backgroundColor:S.b,borderColor:'#fff',borderWidth:2},
 {label:'Port',data:Object.values(D.routes).map(r=>r.port),backgroundColor:S.o,borderColor:'#fff',borderWidth:2},
 {label:'Canal',data:Object.values(D.routes).map(r=>r.canal),backgroundColor:S.g,borderColor:'#fff',borderWidth:2},
 {label:'EU ETS',data:Object.values(D.routes).map(r=>r.ets),backgroundColor:S.y,borderColor:'#fff',borderWidth:2},
 {label:'Misc',data:Object.values(D.routes).map(r=>r.misc),backgroundColor:'#9aa3ad',borderColor:'#fff',borderWidth:2}]},
 options:{maintainAspectRatio:false,scales:{x:{stacked:true,...gridY,ticks:{font:{size:11}}},
 y:{stacked:true,...gridX,ticks:{callback:v=>'$'+(v/1000)+'k'}}},
 plugins:{tooltip:{callbacks:{label:c=>c.dataset.label+': '+usd(c.parsed.y)}}}}});

new Chart(document.getElementById('tce'),{type:'bar',data:{labels:names,datasets:[
 {label:'TCE $/day',data:Object.values(D.routes).map(r=>r.tce),
  backgroundColor:Object.values(D.routes).map(r=>r.tce>=D.benchmark?S.g:S.r),borderRadius:4}]},
 options:{maintainAspectRatio:false,plugins:{legend:{display:false},
  tooltip:{callbacks:{label:c=>usd(c.parsed.y)+'/day'}},
  annotation:false},
  scales:{x:{...gridY,ticks:{font:{size:11}}},y:{...gridX,ticks:{callback:v=>'$'+(v/1000)+'k'}}}}});

const tl=D.tornado.rows.map(r=>r.lever), base=D.tornado.base;
new Chart(document.getElementById('tornado'),{type:'bar',data:{labels:tl,datasets:[
 {label:'Low case',data:D.tornado.rows.map(r=>r.tce_low-base),backgroundColor:S.o,borderRadius:4},
 {label:'High case',data:D.tornado.rows.map(r=>r.tce_high-base),backgroundColor:S.b,borderRadius:4}]},
 options:{indexAxis:'y',maintainAspectRatio:false,
  plugins:{tooltip:{callbacks:{label:c=>c.dataset.label+': '+usd(base+c.parsed.x)+'/day ('+(c.parsed.x>=0?'+':'')+usd(c.parsed.x)+')'}}},
  scales:{x:{...gridX,title:{display:true,text:'change in TCE vs base ('+usd(base)+'/day)',color:INK3},
   ticks:{callback:v=>(v>0?'+':'')+'$'+(v/1000)+'k'}},y:{...gridY}}}});

new Chart(document.getElementById('elas'),{type:'bar',data:{labels:D.elasticity.map(e=>e.lever),datasets:[
 {label:'TCE change for +10%',data:D.elasticity.map(e=>e.d_tce),
  backgroundColor:D.elasticity.map(e=>e.d_tce>=0?S.b:S.o),borderRadius:4}]},
 options:{indexAxis:'y',maintainAspectRatio:false,plugins:{legend:{display:false},
  tooltip:{callbacks:{label:c=>(c.parsed.x>=0?'+':'')+usd(c.parsed.x)+'/day'}}},
  scales:{x:{...gridX,ticks:{callback:v=>(v>0?'+':'')+'$'+(v/1000)+'k'}},y:{...gridY}}}});

function beChart(id,obj){const ks=Object.keys(obj);const cols=[S.b,S.o,S.g];
 new Chart(document.getElementById(id),{type:'line',data:{labels:D.breakeven.brent,
  datasets:ks.map((k,i)=>({label:k,data:obj[k],borderColor:cols[i],backgroundColor:cols[i],
   borderWidth:2,pointRadius:0,pointHoverRadius:5,tension:0}))},
  options:{maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
   plugins:{tooltip:{callbacks:{label:c=>c.dataset.label+': $'+c.parsed.y.toFixed(2)+'/t'}}},
   scales:{x:{...gridY,title:{display:true,text:'Brent $/bbl',color:INK3}},
    y:{...gridX,title:{display:true,text:'break-even fronthaul rate $/t',color:INK3},
     ticks:{callback:v=>'$'+v}}}}});}
beChart('bePort',D.breakeven.port); beChart('beUtil',D.breakeven.util);

new Chart(document.getElementById('speed'),{type:'line',data:{labels:D.speed.speeds,
 datasets:Object.keys(D.speed.curves).map((k,i)=>({label:k,data:D.speed.curves[k],
  borderColor:[S.g,S.b,S.o][i],backgroundColor:[S.g,S.b,S.o][i],borderWidth:2,pointRadius:3,tension:.25}))},
 options:{maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
  plugins:{tooltip:{callbacks:{label:c=>c.dataset.label+': '+usd(c.parsed.y)+'/day'}}},
  scales:{x:{...gridY,title:{display:true,text:'service speed (knots)',color:INK3}},
   y:{...gridX,ticks:{callback:v=>'$'+(v/1000)+'k'}}}}});
</script>
</body></html>
"""


def heatmap():
    g = R["grid"]
    vals = [v for row in g["tce"] for v in row]
    lo, hi = min(vals), max(vals)
    out = ['<table class="hm"><thead><tr><th>Freight level</th>'
           + "".join(f"<th>Brent ${b}</th>" for b in g["brent"]) + "</tr></thead><tbody>"]
    for fm, row in zip(g["freight_mult"], g["tce"]):
        out.append(f"<tr><td>{fm:.2f}× base rate</td>")
        for v in row:
            t = (v - lo) / (hi - lo)
            # single-hue sequential ramp (blue), light -> dark
            r, gg, b = int(247 - 205 * t), int(251 - 131 * t), int(255 - 81 * t)
            ink = "#ffffff" if t > 0.62 else "#0b0b0b"
            mark = "" if v >= A["daily_cost_benchmark"] else " ✕"
            out.append(f'<td style="background:rgb({r},{gg},{b});color:{ink}">${v:,.0f}{mark}</td>')
        out.append("</tr>")
    out.append("</tbody></table>"
               '<p class="note">✕ marks scenarios where TCE falls below the daily vessel-cost benchmark.</p>')
    return "".join(out)


def route_table():
    rows = []
    for k, v in R["route_pnl"].items():
        rows.append(
            f"<tr><td>{v['route']}</td><td class='num'>{v['days']:.1f}</td>"
            f"<td class='num'>{money(v['gross'])}</td><td class='num'>{money(v['bunkers'])}</td>"
            f"<td class='num'>{money(v['port'])}</td><td class='num'>{money(v['canal'])}</td>"
            f"<td class='num'>{money(v['ets'])}</td><td class='num'><strong>{money(v['tce'])}</strong></td>"
            f"<td class='num'>${v['be_rate_fh']:.2f}</td></tr>")
    return "".join(rows)


def tor_table():
    fmt = {"freight_mult": lambda v: f"{v:.2f}×", "brent_usd_bbl": lambda v: f"${v:.2f}",
           "port_time_mult": lambda v: f"{v:.2f}×", "util_fh": lambda v: f"{v:.0%}",
           "util_bh": lambda v: f"{v:.0%}", "speed_kn": lambda v: f"{v:.1f} kn",
           "port_cost_mult": lambda v: f"{v:.2f}×", "eua_eur": lambda v: f"€{v:.0f}"}
    rows = []
    for r in TOR["rows"]:
        f = fmt[r["key"]]
        rows.append(f"<tr><td>{r['lever']}</td><td class='num'>{f(r['low'])}</td><td class='num'>{f(r['high'])}</td>"
                    f"<td class='num'>{money(r['tce_low'])}</td><td class='num'>{money(r['tce_high'])}</td>"
                    f"<td class='num'><strong>{money(r['swing'])}</strong></td></tr>")
    return "".join(rows)


def assumption_table():
    rows = [
        ("Vessel", "38,000 dwt IMO II stainless parcel tanker, 36,000 t usable cargo, 23,500 GT, 600 ft LOA",
         "Generic deep-sea parcel tanker, not a specific ship. GT and LOA drive the port tariffs below"),
        ("Service speed / consumption", "13.5 kn; 27 t/day laden, 24 t/day ballast, 6 t/day in port",
         "Non-eco design of this class; port figure covers pumping, heating and hotel load"),
        ("Bunker prices", f"VLSFO = Brent × 7.7 (${A['brent_usd_bbl']*7.7:,.0f}/t); MGO = Brent × 9.7 (${A['brent_usd_bbl']*9.7:,.0f}/t)",
         "Estimate: ~7.3 bbl/t plus a product premium. Spot-check against Ship &amp; Bunker"),
        ("Brent base case", f"${A['brent_usd_bbl']:.2f}/bbl (trailing 12-month average)",
         "FRED MCOILBRENTEU — data"),
        ("Freight rates (USG–ARA)", f"${R['routes']['1']['rate_fh']:.0f}/t fronthaul, ${R['routes']['1']['rate_bh']:.0f}/t backhaul",
         "ASSUMPTION, calibrated to a plausible TCE. No public chemical rate series exists"),
        ("Utilization", "85% fronthaul, 50% backhaul", "Assumption; the key commercial lever"),
        ("Port time", "5.0 days per operation US Gulf, 3.5 ARA, 3.5 Far East, 4.0 India (incl. waiting)",
         "Estimate. A backhaul cargo adds one operation at each end"),
        ("Port cost, US Gulf", "$813.40 harbour fee + $11.24/ft/day dockage (600&nbsp;ft, tapered) + $45k ancillaries per call",
         "PUBLISHED: Port Houston Tariff No. 14, Subrules 095-096, eff. 1-Jan-2025. Ancillaries (pilotage, tugs, linesmen, agency, private terminal) are an estimate"),
        ("Port cost, ARA", "GT &times; &euro;0.303 + cargo t &times; &euro;0.576 (capped at GT &times; 133.3% &times; rate) + GT &times; &euro;0.065, indexed 3.5%, + $35k ancillaries",
         "PUBLISHED: Port of Rotterdam seaport dues, Tables 1-3 and the chemical-tanker efficiency cap; 2026 indexation per the port. Ancillaries estimated"),
        ("Port cost, Far East / India", "$65k and $55k per call",
         "ESTIMATE — no published schedule transcribed for these ranges"),
        ("Canal tolls", f"Panama ${R['panama_toll_laden']:,.0f} laden / ${R['panama_toll_ballast']:,.0f} ballast; "
                        f"Suez ${R['suez_toll_laden']:,.0f} / ${R['suez_toll_ballast']:,.0f} per transit",
         "Built up from published ACP and SCA schedules; tonnages and ancillaries estimated"),
        ("EU ETS", f"€{A['eua_eur']:.0f}/t CO₂, 100% phase-in, 50% of extra-EU voyage emissions",
         "Allowance price is an estimate — check ICE EUA futures"),
        ("Commissions", "2.5% of gross freight", "Market standard: 1.25% address + 1.25% brokerage"),
        ("Daily vessel cost benchmark", f"${A['daily_cost_benchmark']:,}/day",
         "Estimate: OPEX ~$8.5k + capital recovery ~$9.5k. Substitute your own T/C-in hire"),
    ]
    return "".join(f"<tr><td>{a}</td><td>{b}</td><td>{c}</td></tr>" for a, b, c in rows)


def main():
    swing = {r["lever"]: r["swing"] for r in TOR["rows"]}
    ela = {e["lever"]: e["d_tce"] for e in ELA}
    bunk_share = P1["bunkers"] / P1["voy_costs"] * 100
    port_share = P1["port_days"] / P1["days"] * 100
    be_norm = R["breakeven"]["port"]["1.00x port time"][R["breakeven"]["brent"].index(80)]
    be_sev = R["breakeven"]["port"]["2.00x port time"][R["breakeven"]["brent"].index(80)]
    sp = R["speed"]
    opt = sp["speeds"][max(range(len(sp["speeds"])), key=lambda i: sp["curves"]["Brent $82"][i])]
    sub = {
        "__DATA__": PAY,
        "__BRENT__": f"{A['brent_usd_bbl']:.2f}",
        "__TCE__": money(P1["tce"]), "__DAYS__": f"{P1['days']:.1f}",
        "__GROSS__": f"${P1['gross']/1e6:.2f}m", "__COSTS__": f"${P1['voy_costs']/1e6:.2f}m",
        "__BE__": f"${P1['be_rate_fh']:.2f}",
        "__SW1__": money(swing["Freight rate (market)"]), "__SW2__": money(swing["Backhaul utilization"]),
        "__SWB__": money(swing["Bunker price (Brent)"]),
        "__BP10__": f"{RG['brent_p10']:.0f}", "__BP90__": f"{RG['brent_p90']:.0f}",
        "__BUNKSHARE__": f"{bunk_share:.0f}", "__PORTSHARE__": f"{port_share:.0f}",
        "__ELA_UTIL__": money(ela["Fronthaul utilization"]),
        "__UTIL_VS_BUNKER__": f"{abs(ela['Fronthaul utilization']/ela['Bunker price']):.1f}×",
        "__PORTSWING__": money(swing["Port time (congestion)"]),
        "__OPTSPEED__": f"{opt:.1f}",
        "__OPTSPEED_LOW__": f"{sp['speeds'][max(range(len(sp['curves']['Brent $65'])), key=lambda i: sp['curves']['Brent $65'][i])]:.1f}",
        "__OPTSPEED_HIGH__": f"{sp['speeds'][max(range(len(sp['curves']['Brent $104'])), key=lambda i: sp['curves']['Brent $104'][i])]:.1f}",
        "__WARRISK__": money(R["route_pnl_no_war_risk"]["3"]["tce"] - R["route_pnl"]["3"]["tce"]),
        "__SUEZGAP__": money(R["route_pnl_no_war_risk"]["3"]["tce"] - R["route_pnl"]["4"]["tce"]),
        "__BENCH__": money(A["daily_cost_benchmark"]),
        "__BE_NORMAL__": f"${be_norm:.2f}", "__BE_SEVERE__": f"${be_sev:.2f}",
        "__BE_UPLIFT__": f"{(be_sev/be_norm-1)*100:.0f}",
        "__ROUTETABLE__": route_table(), "__TORTABLE__": tor_table(),
        "__HEATMAP__": heatmap(), "__ASSUMPTIONS__": assumption_table(),
    }
    html = HTML
    for k, v in sub.items():
        html = html.replace(k, str(v))
    (HERE / "index.html").write_text(html)
    print("wrote index.html", len(html), "bytes; opt speed", opt, "be", be_norm, be_sev)


if __name__ == "__main__":
    main()
