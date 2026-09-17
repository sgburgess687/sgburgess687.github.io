"""
Chemical Tanker Voyage Economics Model — analysis.py
=====================================================
Independent analysis using publicly available data. Not affiliated with or
endorsed by any organization named.

What this script does
---------------------
1. Loads the two public time series in ./data (FRED monthly Brent, BLS PPI deep-sea freight)
   and derives the *plausible ranges* used in the sensitivity analysis.
2. Implements the round-voyage Time Charter Equivalent (TCE) engine — the same
   logic as the Excel model (model/chemical_tanker_voyage_model.xlsx) — so every
   number in the write-up is reproducible from code.
3. Runs: route comparison, tornado (plausible ranges), normalized elasticities
   (per-10%-move), break-even freight curves, and a freight x bunker TCE grid.
4. Writes results.json (consumed by index.html and the memo builder).

Everything that is an ASSUMPTION rather than data is in the ASSUMPTIONS dict
below, with a note. Run:  python analysis.py
"""
from __future__ import annotations

import copy
import json
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"

# ---------------------------------------------------------------------------
# 1. ASSUMPTIONS  (all estimates unless marked "tariff" / "data")
# ---------------------------------------------------------------------------
ASSUMPTIONS = {
    # --- Vessel: generic 38,000 dwt IMO II stainless-steel parcel chemical tanker (not a specific ship)
    "dwt": 38_000,                 # t
    "cargo_capacity_t": 36_000,    # t usable cargo (dwt less bunkers, water, stores, constant)
    "pc_ums_t": 20_000,            # Panama Canal net tonnage (estimate for this size)
    "scnt_t": 20_000,              # Suez Canal net tonnage (estimate)
    "hull_value_usd": 35_000_000,  # for war-risk additional premium (estimate)
    "ref_speed_kn": 13.5,          # service speed at which consumption is quoted
    "speed_kn": 13.5,              # operating speed (lever)
    "cons_laden_tpd": 27.0,        # VLSFO/MGO t per sea day at ref speed, laden
    "cons_ballast_tpd": 24.0,      # t per sea day at ref speed, ballast
    "cons_port_tpd": 6.0,          # MGO t per port/canal day (pumping, heating, hotel)
    "sea_margin": 0.05,            # weather/routeing allowance on sea distance
    "speed_exponent": 3.0,         # admiralty (cube) law for consumption vs speed
    # --- Fuel prices: Brent-linked (data), converted with calibrated multipliers (estimate)
    "brent_usd_bbl": 82.01,        # base = trailing-12-month FRED avg (Sep-25..Aug-26)
    "vlsfo_per_brent": 7.7,        # VLSFO $/t per $1/bbl Brent  (=> $631/t at $82)
    "mgo_per_brent": 9.7,          # MGO   $/t per $1/bbl Brent  (=> $795/t at $82)
    # --- Commercial
    "util_fh": 0.85,               # fronthaul cargo utilization (share of capacity)
    "util_bh": 0.50,               # backhaul cargo utilization
    "commission": 0.025,           # 1.25% address + 1.25% brokerage on gross freight
    "misc_voyage_usd": 25_000,     # tank cleaning, surveys, nitrogen, sundries per round voyage
    "port_time_mult": 1.0,         # multiplier on all port days (congestion lever)
    "port_cost_mult": 1.0,         # multiplier on all port disbursements
    "freight_mult": 1.0,           # multiplier on route freight rates (market lever)
    "daily_cost_benchmark": 18_000,  # $/day: OPEX (~$8.5k) + capital recovery (~$9.5k), or T/C-in hire
    # --- EU ETS (2026 = 100% phase-in; 50% of emissions on extra-EU voyages, 100% at EU berth)
    "eua_eur": 75.0,               # EUR/t CO2 (estimate — check ICE EUA futures)
    "usd_per_eur": 1.10,           # estimate
    "ets_phase_in": 1.00,
    "co2_vlsfo": 3.114,            # t CO2 / t fuel (IMO/EU MRV emission factor)
    "co2_mgo": 3.206,
    # --- Canals
    "sdr_usd": 1.35,               # USD per SDR (estimate — check IMF daily rate)
    "suez_surcharge_laden": 0.20,  # chemical tanker surcharge (SCA Circular 7/2023, as reported)
    "suez_surcharge_ballast": 0.15,
    "suez_ancillary_usd": 40_000,  # pilotage, tugs, mooring, agency per Suez transit (estimate)
    "war_risk_ap_pct": 0.0025,     # Red Sea additional premium, % of hull value per transit (illustrative)
    "panama_linehandlers": 8,      # estimate
    # --- Port tariffs: PUBLISHED RATES (Port Houston Tariff 14; Port of Rotterdam 2025 tariffs)
    "gt": 23_500,                  # gross tonnage (estimate for class; PoR's own worked example uses a 23,230 GT chemical tanker)
    "loa_ft": 600,                 # length overall, feet (~183 m) -> Port Houston dockage band 600-650 ft
    "ph_harbor_fee": 813.40,       # $ per vessel entry, >=250 ft — Port Houston Tariff 14, Subrule 096 (eff. 1-Jan-2025)
    "ph_dockage_per_ft": 11.24,    # $ per LOA foot per 24h, 600-650 ft band — Tariff 14, Subrule 095
    "ph_ancillary": 45_000,        # ESTIMATE: pilotage in/out, tugs, linesmen, agency, private-terminal charges
    "usg_berth_share": 0.60,       # share of US Gulf port days actually alongside (rest is waiting) — estimate
    "rtm_gt_rate": 0.303,          # EUR per GT, chemical/gas tanker — PoR Port Tariffs, Table 1
    "rtm_sust_rate": 0.065,        # EUR per GT, sustainability component — PoR Table 3
    "rtm_cargo_rate": 0.576,       # EUR per tonne, "other liquid bulk" — PoR Table 2
    "rtm_cap_pct": 1.333,          # cargo-component cap: GT x 133.3% x cargo rate — PoR efficiency discount, chemical tankers
    "rtm_index": 1.035,            # PoR published 2026 indexation of 3.5% applied to the 2025 schedule
    "rtm_ancillary": 35_000,       # ESTIMATE: pilotage, tugs, boatmen, agency, terminal charges
}

# Port regions: days per operation (load OR discharge, incl. typical waiting) and cost per call.
# All ESTIMATES — replace with agency proforma DAs / published tariffs.
REGIONS = {
    "USG": {"name": "US Gulf (Houston)", "days": 5.0, "eu": 0, "tariff": "PH"},
    "ARA": {"name": "Amsterdam-Rotterdam-Antwerp", "days": 3.5, "eu": 1, "tariff": "RTM"},
    "FEA": {"name": "Far East Asia (Korea/China)", "days": 3.5, "eu": 0, "tariff": "EST", "fixed": 65_000},
    "IND": {"name": "India West Coast", "days": 4.0, "eu": 0, "tariff": "EST", "fixed": 55_000},
}


def dockage_factor(berth_days: float) -> float:
    """Port Houston Tariff 14, Subrule 095: day 1 and 2 at full rate, then 90%, 75%, 60%, 50% thereafter."""
    steps = [1.0, 1.0, 0.90, 0.75, 0.60]
    f, left = 0.0, berth_days
    for st in steps:
        take = min(left, 1.0)
        if take <= 0:
            return f
        f += take * st
        left -= take
    return f + left * 0.50


def port_cost_per_op(code: str, cargo_t: float, a: dict) -> float:
    """Cost of one port operation (a load or a discharge). Published tariff where one exists."""
    g = REGIONS[code]
    if g["tariff"] == "PH":
        berth_days = g["days"] * a["usg_berth_share"] * a["port_time_mult"]
        dockage = a["loa_ft"] * a["ph_dockage_per_ft"] * dockage_factor(berth_days)
        return a["ph_harbor_fee"] + dockage + a["ph_ancillary"]
    if g["tariff"] == "RTM":
        eur = a["gt"] * (a["rtm_gt_rate"] + a["rtm_sust_rate"])
        eur += min(cargo_t * a["rtm_cargo_rate"], a["gt"] * a["rtm_cap_pct"] * a["rtm_cargo_rate"])
        return eur * a["rtm_index"] * a["usd_per_eur"] + a["rtm_ancillary"]
    return g["fixed"]


# Routes: fronthaul A->B laden, backhaul B->A. Distances are APPROXIMATE (verify with a distance table).
# Freight rates ($/t) are CALIBRATED ASSUMPTIONS, not observed fixtures.
ROUTES = {
    1: {"name": "USG-ARA transatlantic", "A": "USG", "B": "ARA", "nm": 5_000, "eca_nm": 650,
        "canal": "None", "canal_days": 0.0, "eu_share": 0.5, "rate_fh": 52.0, "rate_bh": 38.0},
    2: {"name": "USG-Far East via Panama", "A": "USG", "B": "FEA", "nm": 9_650, "eca_nm": 200,
        "canal": "Panama", "canal_days": 1.5, "eu_share": 0.0, "rate_fh": 92.0, "rate_bh": 52.0},
    3: {"name": "ARA-India WC via Suez", "A": "ARA", "B": "IND", "nm": 6_350, "eca_nm": 450,
        "canal": "Suez", "canal_days": 1.0, "eu_share": 0.5, "rate_fh": 75.0, "rate_bh": 55.0},
    4: {"name": "ARA-India WC via Cape", "A": "ARA", "B": "IND", "nm": 10_850, "eca_nm": 450,
        "canal": "None", "canal_days": 0.0, "eu_share": 0.5, "rate_fh": 75.0, "rate_bh": 55.0},
}

# Panama tariff (ACP schedule, extracted in tanker_voyage_data/canal_tolls) — Super vessel, chemical tanker
PANAMA = {"fixed": 100_000, "capacity_rate": 4.00, "ballast_factor": 0.85, "tug": 7_000,
          "linehandler": 270, "reservation": 10_500, "security": 1_250, "fresh_water": 10_000,
          "pilotage": 1_275 * 2}
# Suez petroleum-product-tanker bands (SDR/SCNT), SCA Circular 7/2023
SUEZ_BANDS = [  # (band size, laden, ballast)
    (5_000, 11.04, 9.40), (5_000, 7.82, 6.64), (10_000, 5.91, 5.04), (20_000, 3.93, 2.50),
    (30_000, 3.84, 2.14), (50_000, 3.46, 1.85), (10**9, 3.34, 1.82)]


def panama_toll(a, laden: bool) -> float:
    toll = PANAMA["fixed"] + PANAMA["capacity_rate"] * a["pc_ums_t"]
    if not laden:
        toll *= PANAMA["ballast_factor"]
    anc = (PANAMA["tug"] + PANAMA["linehandler"] * a["panama_linehandlers"] + PANAMA["reservation"]
           + PANAMA["security"] + PANAMA["fresh_water"] + PANAMA["pilotage"])
    return toll + anc


def suez_toll(a, laden: bool) -> float:
    remaining, sdr = a["scnt_t"], 0.0
    for size, l, b in SUEZ_BANDS:
        t = min(remaining, size)
        sdr += t * (l if laden else b)
        remaining -= t
        if remaining <= 0:
            break
    sur = a["suez_surcharge_laden"] if laden else a["suez_surcharge_ballast"]
    return sdr * (1 + sur) * a["sdr_usd"] + a["suez_ancillary_usd"] + a["war_risk_ap_pct"] * a["hull_value_usd"]


def voyage(route_id: int, a: dict | None = None) -> dict:
    """Round voyage A->B (fronthaul) and B->A (backhaul). Returns full P&L build-up."""
    a = a or ASSUMPTIONS
    r = ROUTES[route_id]
    A, B = REGIONS[r["A"]], REGIONS[r["B"]]
    bh_laden = a["util_bh"] > 0
    vlsfo_px = a["brent_usd_bbl"] * a["vlsfo_per_brent"]
    mgo_px = a["brent_usd_bbl"] * a["mgo_per_brent"]
    spd = a["speed_kn"]
    cube = (spd / a["ref_speed_kn"]) ** a["speed_exponent"]

    sea_days_leg = r["nm"] * (1 + a["sea_margin"]) / (spd * 24)
    eca_frac = r["eca_nm"] / r["nm"]
    cons_fh = a["cons_laden_tpd"] * cube
    cons_bh = (a["cons_laden_tpd"] if bh_laden else a["cons_ballast_tpd"]) * cube
    sea_fuel = {"fh": sea_days_leg * cons_fh, "bh": sea_days_leg * cons_bh}
    vlsfo_t = (sea_fuel["fh"] + sea_fuel["bh"]) * (1 - eca_frac)
    mgo_sea_t = (sea_fuel["fh"] + sea_fuel["bh"]) * eca_frac

    ops = 2 if bh_laden else 1  # port operations per region
    port_days_A = A["days"] * ops * a["port_time_mult"]
    port_days_B = B["days"] * ops * a["port_time_mult"]
    port_days = port_days_A + port_days_B
    canal_days = r["canal_days"] * 2
    mgo_port_t = (port_days + canal_days) * a["cons_port_tpd"]

    bunker_cost = vlsfo_t * vlsfo_px + (mgo_sea_t + mgo_port_t) * mgo_px
    qty_fh_t = a["cargo_capacity_t"] * a["util_fh"]
    qty_bh_t = a["cargo_capacity_t"] * a["util_bh"]
    # A: loads the fronthaul cargo, and (if there is one) discharges the backhaul cargo
    port_cost = (port_cost_per_op(r["A"], qty_fh_t, a) + port_cost_per_op(r["B"], qty_fh_t, a))
    if bh_laden:
        port_cost += (port_cost_per_op(r["B"], qty_bh_t, a) + port_cost_per_op(r["A"], qty_bh_t, a))
    port_cost *= a["port_cost_mult"]

    if r["canal"] == "Panama":
        canal_cost = panama_toll(a, True) + panama_toll(a, bh_laden)
    elif r["canal"] == "Suez":
        canal_cost = suez_toll(a, True) + suez_toll(a, bh_laden)
    else:
        canal_cost = 0.0

    # EU ETS
    co2_sea = (vlsfo_t * a["co2_vlsfo"] + mgo_sea_t * a["co2_mgo"])
    co2_eu_port = (port_days_A * A["eu"] + port_days_B * B["eu"]) * a["cons_port_tpd"] * a["co2_mgo"]
    ets_t = (co2_sea * r["eu_share"] + co2_eu_port) * a["ets_phase_in"]
    ets_cost = ets_t * a["eua_eur"] * a["usd_per_eur"]

    qty_fh, qty_bh = qty_fh_t, qty_bh_t
    rate_fh = r["rate_fh"] * a["freight_mult"]
    rate_bh = r["rate_bh"] * a["freight_mult"]
    gross = qty_fh * rate_fh + qty_bh * rate_bh
    comm = gross * a["commission"]
    voy_costs = bunker_cost + port_cost + canal_cost + ets_cost + a["misc_voyage_usd"]
    net = gross - comm - voy_costs
    days = sea_days_leg * 2 + port_days + canal_days
    tce = net / days
    result = net - a["daily_cost_benchmark"] * days
    # break-even fronthaul rate (TCE == benchmark)
    be_rate_fh = ((a["daily_cost_benchmark"] * days + voy_costs) / (1 - a["commission"])
                  - qty_bh * rate_bh) / qty_fh
    return dict(route=r["name"], days=days, sea_days=sea_days_leg * 2, port_days=port_days,
                canal_days=canal_days, vlsfo_t=vlsfo_t, mgo_t=mgo_sea_t + mgo_port_t,
                vlsfo_px=vlsfo_px, mgo_px=mgo_px, gross=gross, commission=comm,
                bunkers=bunker_cost, port=port_cost, canal=canal_cost, ets=ets_cost,
                misc=a["misc_voyage_usd"], voy_costs=voy_costs, net=net, tce=tce,
                result=result, be_rate_fh=be_rate_fh, qty_fh=qty_fh, qty_bh=qty_bh,
                rate_fh=rate_fh, rate_bh=rate_bh, co2_ets_t=ets_t)


def with_(**kw):
    a = copy.deepcopy(ASSUMPTIONS)
    a.update(kw)
    return a


# ---------------------------------------------------------------------------
# 2. DATA -> plausible ranges
# ---------------------------------------------------------------------------
def load_data():
    brent = pd.read_csv(DATA / "fred_brent_monthly_MCOILBRENTEU.csv", parse_dates=["date"]).set_index("date")["brent_usd_bbl"]
    ppi = pd.read_csv(DATA / "bls_ppi_deep_sea_freight_WPU30130101.csv", parse_dates=["date"]).set_index("date")["value"]
    return brent, ppi


def ranges(brent, ppi):
    b5 = brent["2021-09":"2026-08"]
    yoy = ppi.pct_change(12).dropna()["2016":]
    return {
        "brent_latest": float(brent.iloc[-1]), "brent_latest_month": brent.index[-1].strftime("%b-%Y"),
        "brent_ttm": round(float(brent.iloc[-12:].mean()), 2),
        "brent_5y": round(float(b5.mean()), 2),
        "brent_p10": round(float(b5.quantile(0.10)), 2), "brent_p90": round(float(b5.quantile(0.90)), 2),
        "brent_min5": float(b5.min()), "brent_max5": float(b5.max()),
        "ppi_yoy_p10": round(float(yoy.quantile(0.10)), 3), "ppi_yoy_p90": round(float(yoy.quantile(0.90)), 3),
        "ppi_latest": float(ppi.iloc[-1]), "ppi_latest_month": ppi.index[-1].strftime("%b-%Y"),
    }


# Tornado levers: (label, key, low, high, unit label). Freight/Brent bands from data (rounded).
def levers(rg):
    return [
        ("Freight rate (market)", "freight_mult", 1 + round(rg["ppi_yoy_p10"], 2), 1 + round(rg["ppi_yoy_p90"], 2), "x base rate"),
        ("Bunker price (Brent)", "brent_usd_bbl", rg["brent_p10"], rg["brent_p90"], "$/bbl"),
        ("Port time (congestion)", "port_time_mult", 0.75, 1.75, "x base port days"),
        ("Fronthaul utilization", "util_fh", 0.70, 0.95, "share of capacity"),
        ("Backhaul utilization", "util_bh", 0.25, 0.75, "share of capacity"),
        ("Service speed", "speed_kn", 12.0, 14.5, "knots"),
        ("Port disbursements", "port_cost_mult", 0.75, 1.25, "x base"),
        ("EU carbon price (EUA)", "eua_eur", 50.0, 110.0, "EUR/t"),
    ]


def tornado(route_id, rg):
    base = voyage(route_id)["tce"]
    rows = []
    for label, key, lo, hi, unit in levers(rg):
        t_lo = voyage(route_id, with_(**{key: lo}))["tce"]
        t_hi = voyage(route_id, with_(**{key: hi}))["tce"]
        rows.append(dict(lever=label, key=key, low=lo, high=hi, unit=unit,
                         tce_low=t_lo, tce_high=t_hi, swing=abs(t_hi - t_lo)))
    rows.sort(key=lambda r: -r["swing"])
    return base, rows


def elasticities(route_id):
    """TCE change ($/day) for a +10% move in each lever (apples-to-apples leverage)."""
    base = voyage(route_id)["tce"]
    out = []
    for label, key in [("Freight rate", "freight_mult"), ("Bunker price", "brent_usd_bbl"),
                       ("Port time", "port_time_mult"), ("Fronthaul utilization", "util_fh"),
                       ("Backhaul utilization", "util_bh"), ("Service speed", "speed_kn"),
                       ("Port disbursements", "port_cost_mult")]:
        v = ASSUMPTIONS[key] * 1.10
        out.append(dict(lever=label, d_tce=voyage(route_id, with_(**{key: v}))["tce"] - base))
    return sorted(out, key=lambda r: -abs(r["d_tce"]))


def breakeven_curves(route_id, brents, port_mults=(1.0, 1.5, 2.0), utils=(0.70, 0.85, 0.95)):
    port = {f"{m:.2f}x port time": [voyage(route_id, with_(brent_usd_bbl=b, port_time_mult=m))["be_rate_fh"] for b in brents]
            for m in port_mults}
    util = {f"{u:.0%} utilization": [voyage(route_id, with_(brent_usd_bbl=b, util_fh=u))["be_rate_fh"] for b in brents]
            for u in utils}
    return port, util


def grid(route_id, freight_mults, brents):
    return [[voyage(route_id, with_(freight_mult=f, brent_usd_bbl=b))["tce"] for b in brents] for f in freight_mults]


def speed_curve(route_id, brents=(65, 82, 104), speeds=(10.0, 10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 14.5, 15.0)):
    return {f"Brent ${b}": [voyage(route_id, with_(brent_usd_bbl=b, speed_kn=s))["tce"] for s in speeds] for b in brents}, list(speeds)


def main():
    brent, ppi = load_data()
    rg = ranges(brent, ppi)
    res = {"ranges": rg, "assumptions": ASSUMPTIONS, "regions": REGIONS, "routes": ROUTES,
           "panama_toll_laden": panama_toll(ASSUMPTIONS, True), "panama_toll_ballast": panama_toll(ASSUMPTIONS, False),
           "suez_toll_laden": suez_toll(ASSUMPTIONS, True), "suez_toll_ballast": suez_toll(ASSUMPTIONS, False)}
    res["route_pnl"] = {rid: voyage(rid) for rid in ROUTES}
    res["route_pnl_no_backhaul"] = {rid: voyage(rid, with_(util_bh=0.0)) for rid in ROUTES}
    res["route_pnl_no_war_risk"] = {rid: voyage(rid, with_(war_risk_ap_pct=0.0)) for rid in (3, 4)}
    res["tornado"] = {}
    res["elasticity"] = {}
    for rid in ROUTES:
        base, rows = tornado(rid, rg)
        res["tornado"][rid] = {"base": base, "rows": rows}
        res["elasticity"][rid] = elasticities(rid)
    brents = list(range(50, 135, 5))
    pc, uc = breakeven_curves(1, brents)
    res["breakeven"] = {"brent": brents, "port": pc, "util": uc,
                        "by_route": {ROUTES[r]["name"]: [voyage(r, with_(brent_usd_bbl=b))["be_rate_fh"] for b in brents] for r in ROUTES}}
    fm = [0.8, 0.9, 1.0, 1.1, 1.2, 1.3]
    gb = [60, 70, 80, 90, 100, 110, 120]
    res["grid"] = {"freight_mult": fm, "brent": gb, "tce": grid(1, fm, gb)}
    sc, sp = speed_curve(1)
    res["speed"] = {"speeds": sp, "curves": sc}
    res["brent_series"] = {"dates": [d.strftime("%Y-%m") for d in brent.index], "values": brent.round(2).tolist()}
    res["ppi_series"] = {"dates": [d.strftime("%Y-%m") for d in ppi.index], "values": ppi.round(2).tolist()}
    (HERE / "results.json").write_text(json.dumps(res, indent=1, default=float))

    # console summary
    print("Ranges:", rg)
    for rid, p in res["route_pnl"].items():
        print(f"R{rid} {p['route']:<26} days {p['days']:5.1f}  gross ${p['gross']/1e6:5.2f}m  "
              f"bunkers ${p['bunkers']/1e3:6.0f}k port ${p['port']/1e3:4.0f}k canal ${p['canal']/1e3:4.0f}k "
              f"ETS ${p['ets']/1e3:4.0f}k  TCE ${p['tce']:,.0f}/d  result ${p['result']/1e3:,.0f}k  BE ${p['be_rate_fh']:.2f}/t")
    for r in res["tornado"][1]["rows"]:
        print(f"  {r['lever']:<26} {r['low']:>7} {r['high']:>7}  {r['tce_low']:>9,.0f} {r['tce_high']:>9,.0f} swing {r['swing']:,.0f}")
    for e in res["elasticity"][1]:
        print(f"  +10% {e['lever']:<24} {e['d_tce']:+,.0f} $/day")
    return res


if __name__ == "__main__":
    main()
