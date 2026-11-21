#!/usr/bin/env python3
"""
Synthetic sample-data generator for ANDON.

Produces a weekly supply-planning workbook with the exact layout ANDON reads
(sheets "Need", "New R" and "Open Orders": same header rows, same column names,
same number formats), filled with entirely fictional data:

  * part numbers, scheduling-agreement numbers, suppliers, contacts, projects
    and planner names are invented;
  * quantities, prices, stocks and delivery schedules are drawn from random
    distributions.

The layout and the orders of magnitude (number of parts, share of each risk
category, spread of prices and volumes) are modelled on a real planning
workbook; no real identifier, name or figure is copied.

The data is internally consistent, the way a real planning file is:

    projected stock[week] = projected stock[week-1]
                            - consumption[week]
                            + deliveries landing at the plant that week

with deliveries dated at their arrival at the logistics hub and reaching the
plant one week later, exactly as ANDON assumes.

Usage:
    pip install openpyxl
    python tools/generate_sample_data.py            # writes into sample-data/
    python tools/generate_sample_data.py --seed 7   # another fictional plant
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import math
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Layout constants (they mirror the weekly export ANDON was built for)
# ----------------------------------------------------------------------------
DEFAULT_SEED = 342026
N_WEEKS = 36                                  # consumption horizon and stock horizon
FIRST_NEED_MONDAY = dt.date(2026, 8, 17)      # "CW  34"
FIRST_NEED_WEEK = 34
FIRST_STOCK_WEEK = 33                         # stock table starts one week earlier
OO_WEEK_COLUMNS = 67                          # open-orders table: weeks 33..52 then 1..47
HUB_TO_PLANT_WEEKS = 1

NEED_HEADER_ROW = 23                          # 1-based Excel rows
OO_HEADER_ROW = 17
MASTER_HEADER_ROW = 37

N_RAW = 2400            # raw materials present in the forecast (PN starts with "1")
N_NON_RAW = 1500        # finished / semi-finished goods (ignored by ANDON)
N_NO_FORECAST = 230     # raw materials in the master data without a forecast line
N_NOT_IN_MASTER = 14    # forecast lines missing from the master data
N_Z_CODES = 30          # consumables with alphanumeric codes

EUR_1 = r'_-[$€-2]\ * #,##0.0_-;\-[$€-2]\ * #,##0.0_-;_-[$€-2]\ * "-"??_-;_-@'
EUR_2 = r'_-[$€-2]\ * #,##0.00_-;\-[$€-2]\ * #,##0.00_-;_-[$€-2]\ * "-"??_-;_-@'

OWNERS = [("Salma", 0.29), ("Youssef", 0.26), ("Hicham", 0.23), ("Leila", 0.22)]

# segment -> (master "Status" value, master "Cat" value, project tags)
SEGMENTS = {
    "HeatPump": ("HeatPump", "Heat pump", [
        "HP Aria", "HP Boreal", "HP Cirrus", "HP Delta", "HP Evo", "HP Fjord",
        "HP Geyser", "HP Monoblock", "HP Split 8kW", "HP Split 12kW", "HP Hydro"]),
    "Auto": ("Auto", "Auto", [
        "Seat Comfort", "Door Module", "eDrive M1", "eDrive M2", "BMS Link",
        "Trailer Kit", "Sensor Hub", "Steering Col", "Mirror Fold", "Tailgate Lift",
        "Roof Module", "Charge Port", "ADAS Front", "Cockpit Gen3", "Lighting RL"]),
    "New Project": ("New Project", "New Project", ["NP Falcon", "NP Orion", "NP Lynx", "NP Vega"]),
    "Industrial": ("Industrial", "Industrial", [
        "Pump Drive", "HVAC Control", "Robotics Arm", "Solar Inverter", "Rail Signal"]),
    "Spare Parts": ("Spare Parts", "Spare Part & CS", ["Service Kit", "Aftermarket"]),
}
SEGMENT_WEIGHTS = [("HeatPump", 0.34), ("Auto", 0.36), ("New Project", 0.13),
                   ("Industrial", 0.10), ("Spare Parts", 0.07)]

# family -> (share, median unit price EUR, price sigma, median weekly volume,
#            fractional quantities?, supplier specialty)
FAMILIES = {
    "terminal":   (0.19, 0.04, 0.6, 5000, False, "contact"),
    "seal":       (0.09, 0.02, 0.6, 5500, False, "seal"),
    "housing":    (0.17, 0.30, 0.8, 900,  False, "plastic"),
    "wire":       (0.12, 0.15, 0.7, 3500, True,  "wire"),
    "tube":       (0.06, 0.45, 0.7, 500,  True,  "plastic"),
    "tape":       (0.05, 1.40, 0.5, 90,   False, "chem"),
    "clip":       (0.09, 0.07, 0.7, 1800, False, "plastic"),
    "fuse_relay": (0.05, 0.90, 0.7, 350,  False, "electro"),
    "module":     (0.03, 7.00, 0.8, 120,  False, "electro"),
    "thermal":    (0.06, 5.00, 0.9, 150,  False, "thermal"),
    "label":      (0.03, 0.01, 0.5, 3000, False, "pack"),
    "packaging":  (0.04, 0.90, 0.7, 180,  False, "pack"),
    "chemical":   (0.02, 6.00, 0.6, 10,   True,  "chem"),
}

SPECIALTY_WORDS = {
    "contact": ["Connect", "Terminals", "Contacts", "Interconnect"],
    "wire":    ["Cable", "Wire", "Cables"],
    "plastic": ["Plastics", "Polymers", "Molding"],
    "seal":    ["Seals", "Elastomers", "Sealing"],
    "chem":    ["Adhesives", "Tapes", "Chemicals"],
    "electro": ["Electronics", "Sensors", "Controls"],
    "thermal": ["Thermal", "Valves", "Hydronics"],
    "pack":    ["Packaging", "Labels", "Pack"],
}
# local-language variants, used only for suppliers of the matching countries
LOCAL_WORDS = {
    "DE": {"wire": "Kabel", "plastic": "Kunststoff", "seal": "Dichtungen", "chem": "Chemie",
           "electro": "Elektronik", "thermal": "Klima"},
    "CH": {"wire": "Kabel", "electro": "Elektronik"},
    "IT": {"wire": "Cavi", "plastic": "Plastica", "electro": "Elettronica"},
    "FR": {"wire": "Câbles", "pack": "Emballages", "seal": "Étanchéité"},
    "MA": {"wire": "Câblage", "pack": "Emballages"},
    "TN": {"wire": "Câblage", "pack": "Emballages"},
    "ES": {"wire": "Cables", "plastic": "Plásticos"},
}

# Invented company-name roots: any resemblance to an existing company is accidental.
NAME_ROOTS = """
Velmar Tessaro Orvane Quintrel Bravant Solvique Kardenza Norlith Aurexa Zentaro Halvik
Meridane Lumetro Draxon Ibrano Jalvek Korvent Lysando Movanti Ostrel Pravik Rundal Sivera
Tulmann Ulvane Vexoria Wendrik Xaltis Yorvane Zemra Albrecon Bexora Calvoro Delmaro Eskaro
Fintrava Galvesto Hestrel Indrano Jorvell Kelstra Lorvanti Mastrel Nevaro Olbrex Pentaro
Questra Rivanto Sorbex Talvion Umbrel Varnex Welstra Yalkor Zorvane Altemis Brevia Corvane
Dunmar Elstrom Farvent Gorland Helmar Istrana Juvelo Kromar Lindroz Monvert Nordane Orlitz
Polmare Rostane Sandrel Tirvane Ulmaro Vandrel
""".split()

# country code -> (weight, legal forms, cities, incoterms, lead-time weeks, hub transit days, sourcing tag)
COUNTRIES = {
    "DE": (0.20, ["GmbH", "AG", "GmbH & Co. KG"], ["Stuttgart", "Nürnberg", "Wuppertal", "Regensburg"], ["FCA", "DAP"], (4, 8), 14, "Inter"),
    "IT": (0.10, ["S.p.A.", "S.r.l."], ["Torino", "Bologna", "Brescia"], ["FCA", "DAP"], (4, 7), 14, "Inter"),
    "ES": (0.08, ["S.L.", "S.A."], ["Barcelona", "Zaragoza", "Vigo"], ["FCA", "DAP"], (3, 6), 14, "Inter"),
    "FR": (0.07, ["SAS", "SA"], ["Lyon", "Grenoble", "Lille"], ["FCA", "DAP"], (4, 7), 14, "Inter"),
    "MA": (0.12, ["SARL", "SARL AU", "S.A."], ["Casablanca", "Kénitra", "Mohammedia", "Rabat"], ["DAP", "EXW"], (1, 3), 7, "Nat"),
    "PL": (0.06, ["Sp. z o.o."], ["Wrocław", "Poznań"], ["FCA"], (4, 8), 14, "Inter"),
    "CZ": (0.05, ["s.r.o."], ["Brno", "Plzeň"], ["FCA"], (4, 8), 14, "Inter"),
    "RO": (0.05, ["S.R.L."], ["Timișoara", "Sibiu"], ["FCA"], (4, 8), 14, "Inter"),
    "TR": (0.05, ["A.Ş."], ["Bursa", "İzmir"], ["FCA", "CIP"], (4, 7), 14, "Inter"),
    "TN": (0.04, ["SARL"], ["Sousse", "Bizerte"], ["FCA"], (3, 6), 14, "Inter"),
    "PT": (0.04, ["Lda"], ["Porto", "Braga"], ["FCA"], (3, 6), 14, "Inter"),
    "CN": (0.07, ["Co., Ltd."], ["Suzhou", "Shenzhen", "Ningbo"], ["FOB", "CIF"], (12, 16), 14, "China"),
    "NL": (0.02, ["B.V."], ["Eindhoven"], ["FCA"], (4, 6), 14, "Inter"),
    "HU": (0.02, ["Kft."], ["Győr"], ["FCA"], (4, 8), 14, "Inter"),
    "CH": (0.02, ["AG"], ["Winterthur"], ["FCA"], (4, 8), 14, "Inter"),
    "UK": (0.01, ["Ltd"], ["Birmingham"], ["DAP"], (5, 9), 14, "Inter"),
}
N_SUPPLIERS = 150


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def weighted(rng: random.Random, pairs):
    r, acc = rng.random(), 0.0
    for value, w in pairs:
        acc += w
        if r <= acc:
            return value
    return pairs[-1][0]


def nice_lot(x: float) -> int:
    """Round to a 'catalogue-like' lot size: 1 / 2 / 2.5 / 5 x 10^k."""
    if x <= 1:
        return 1
    k = math.floor(math.log10(x))
    base = 10 ** k
    for m in (1, 2, 2.5, 5, 10):
        if x <= m * base * 1.2:
            return max(1, int(round(m * base)))
    return int(10 * base)


def num(v: float, fractional: bool):
    """Integers stay integers; fractional quantities keep 2 decimals."""
    if fractional:
        r = round(v, 2)
        return int(r) if float(r).is_integer() else r
    return int(round(v))


def week_number(offset_from_33: int) -> int:
    w = FIRST_STOCK_WEEK + offset_from_33
    while w > 52:
        w -= 52
    return w


# ----------------------------------------------------------------------------
# Master entities
# ----------------------------------------------------------------------------
def build_suppliers(rng: random.Random):
    suppliers, used = [], set()
    countries = [(c, v[0]) for c, v in COUNTRIES.items()]
    specialties = list(SPECIALTY_WORDS)
    while len(suppliers) < N_SUPPLIERS:
        root = rng.choice(NAME_ROOTS)
        spec = rng.choice(specialties)
        country = weighted(rng, countries)
        _, forms, cities, incoterms, lt, hub_days, tag = COUNTRIES[country]
        word = rng.choice(SPECIALTY_WORDS[spec])
        if spec in LOCAL_WORDS.get(country, {}) and rng.random() < 0.6:
            word = LOCAL_WORDS[country][spec]
        name = f"{root} {word} {rng.choice(forms)}"
        if name in used:
            continue
        used.add(name)
        slug = "".join(ch for ch in f"{root}-{spec}".lower() if ch.isalnum() or ch == "-")
        suppliers.append({
            "name": name, "specialty": spec, "country": country,
            "location": f"{rng.choice(cities)}, {country}",
            "incoterm": rng.choice(incoterms),
            "lead_time": (lambda n: f"{n} week" + ("s" if n > 1 else ""))(rng.randint(*lt)),
            "hub_days": hub_days, "tag": tag,
            "contact": f"sales@{slug}.example",
            "code": root[:3].upper(),
            "owner": None,                         # set by assign_owners()
            "weight": rng.lognormvariate(0, 0.9),   # a few big suppliers, many small ones
        })
    return suppliers


def assign_owners(suppliers, parts):
    """Each planner owns whole suppliers; portfolios follow the OWNERS shares."""
    load = {}
    for p in parts:
        load[p["supplier"]["name"]] = load.get(p["supplier"]["name"], 0) + 1
    total = sum(load.values()) or 1
    got = {o: 0 for o, _ in OWNERS}
    for s in sorted(suppliers, key=lambda s: -load.get(s["name"], 0)):
        owner = min(OWNERS, key=lambda ow: got[ow[0]] / ow[1])[0]
        s["owner"] = owner
        got[owner] += load.get(s["name"], 0) / total


def pick_supplier(rng, suppliers, specialty):
    pool = [s for s in suppliers if s["specialty"] == specialty] or suppliers
    total = sum(s["weight"] for s in pool)
    r, acc = rng.random() * total, 0.0
    for s in pool:
        acc += s["weight"]
        if r <= acc:
            return s
    return pool[-1]


# ----------------------------------------------------------------------------
# Demand, stock and delivery schedules
# ----------------------------------------------------------------------------
def gen_need(rng, level, pattern, fractional):
    w = [0.0] * N_WEEKS
    if pattern == "steady":
        trend = rng.uniform(-0.25, 0.25)
        for j in range(N_WEEKS):
            w[j] = level * (1 + trend * (j / (N_WEEKS - 1) - 0.5)) * rng.lognormvariate(0, 0.18)
    elif pattern == "lumpy":
        p = rng.uniform(0.15, 0.6)
        lot = level / p
        for j in range(N_WEEKS):
            if rng.random() < p:
                w[j] = lot * rng.choice([0.5, 1, 1, 1, 2]) * rng.lognormvariate(0, 0.10)
    elif pattern == "rampup":
        start, ramp = rng.randint(1, 16), rng.randint(5, 10)
        for j in range(start, N_WEEKS):
            w[j] = level * min(1.0, (j - start + 1) / ramp) * rng.lognormvariate(0, 0.15)
    elif pattern == "phaseout":
        end = rng.randint(8, 28)
        for j in range(end):
            w[j] = level * (1 - j / end) * rng.lognormvariate(0, 0.15)
    elif pattern == "late":
        start = rng.randint(4, 24)
        for j in range(start, N_WEEKS):
            w[j] = level * rng.lognormvariate(0, 0.18)
    # year-end shutdown: CW 52 and CW 01 run at reduced volume
    if rng.random() < 0.85:
        w[52 - FIRST_NEED_WEEK] *= 0.2
        w[53 - FIRST_NEED_WEEK] *= 0.5
    return [num(x, fractional) if x > 0 else 0 for x in w]


def project(stock, raben, need, edi):
    """Stock calendar: index 0 = week 33 baseline, need[j] is consumed at index j+1,
    hub stock lands at index 1, an order dated week 33+k lands at index k+1."""
    proj = [stock]
    for i in range(1, N_WEEKS):
        arrivals = (raben if i == 1 else 0) + (edi.get(i - 1, 0) if i >= 2 else 0)
        proj.append(proj[-1] - need[i - 1] + arrivals)
    return proj


def plan_orders(rng, need, stock, raben, policy, lot, first_landing, last_landing):
    """Greedy replenishment: whenever the projected stock would fall under the
    policy's target at a week where an order may still land, add whole lots."""
    peak12 = max(need[:12]) if need else 0
    overall = sum(need) / len(need)
    edi, level = {}, stock
    for i in range(1, N_WEEKS):
        level = level - need[i - 1] + (raben if i == 1 else 0)
        fwd = need[i:i + 4]
        fwd_avg = (sum(fwd) / len(fwd)) if fwd else overall
        fwd_max2 = max(need[i:i + 2]) if need[i:i + 2] else overall
        if policy == "robust":
            target = max(2.3 * peak12, 2.6 * fwd_avg)
        elif policy == "delay_ok":
            target = 2.6 * max(fwd_max2, fwd_avg)
        elif policy == "jit":
            target = rng.uniform(0.3, 0.9) * max(fwd_avg, overall * 0.5)
        else:  # "late": sensible target, but nothing can land before first_landing
            target = max(2.3 * peak12, 2.6 * fwd_avg)
        if first_landing <= i <= last_landing and level < target:
            qty = math.ceil((target - level) / lot) * lot
            edi[i - 1] = edi.get(i - 1, 0) + qty
            level += qty
    return edi


def build_raw_items(rng, suppliers):
    fam_pairs = [(f, v[0]) for f, v in FAMILIES.items()]
    pns = sorted(rng.sample(range(14_000_001, 14_099_999), N_RAW))
    sa_pool = rng.sample(range(5_600_000_001, 5_600_099_999), N_RAW)
    items = []
    for idx, pn in enumerate(pns):
        fam = weighted(rng, fam_pairs)
        _, p_med, p_sig, vol_med, fractional, spec = FAMILIES[fam]
        seg = weighted(rng, SEGMENT_WEIGHTS)
        sup = pick_supplier(rng, suppliers, spec)
        price = round(rng.lognormvariate(math.log(p_med), p_sig), 4) if rng.random() < 0.78 else 0
        tags = SEGMENTS[seg][2]
        projects = " / ".join(rng.sample(tags, min(len(tags), rng.choice([1, 1, 1, 2, 2, 3, 4]))))
        if rng.random() < 0.02:
            projects = "0"

        r = rng.random()
        arche = ("inactive" if r < 0.19 else "dead_stock" if r < 0.39 else
                 "slow_covered" if r < 0.58 else "flow" if r < 0.85 else "uncovered")
        level = min(200_000, max(0.5, rng.lognormvariate(math.log(vol_med), 1.6)))
        need = [0] * N_WEEKS
        stock = raben = 0
        edi, backlog, wk33 = {}, 0, 0
        has_sa = rng.random() < 0.5
        lot = 0

        if arche == "dead_stock":
            if price > 0:
                stock = max(1, rng.lognormvariate(math.log(800), 1.5) / price)
            else:
                stock = rng.lognormvariate(math.log(vol_med * 2), 1.4)
            has_sa = rng.random() < 0.4
        elif arche == "slow_covered":
            level *= 0.06
            pattern = weighted(rng, [("lumpy", 0.6), ("steady", 0.2), ("phaseout", 0.2)])
            need = gen_need(rng, level, pattern, fractional)
            stock = sum(need) * rng.uniform(1.05, 5) + (level if sum(need) == 0 else 0)
            has_sa = rng.random() < 0.6
        elif arche == "uncovered":
            level *= 0.3
            pattern = weighted(rng, [("steady", 0.4), ("lumpy", 0.3), ("rampup", 0.15), ("late", 0.15)])
            need = gen_need(rng, level, pattern, fractional)
            stock = sum(need) * (rng.uniform(0.03, 0.35) if rng.random() < 0.5 else rng.uniform(0.35, 0.98))
            if rng.random() < 0.12 and level < 400:
                stock = 0
            has_sa = rng.random() < 0.45
        elif arche == "flow":
            pattern = weighted(rng, [("steady", 0.5), ("lumpy", 0.2), ("rampup", 0.12),
                                     ("phaseout", 0.10), ("late", 0.08)])
            need = gen_need(rng, level, pattern, fractional)
            active = [x for x in need if x > 0] or [level]
            typical = sum(active) / len(active)
            lot = nice_lot(typical * rng.choice([1, 2, 2, 3, 4, 6]))
            policy = weighted(rng, [("robust", 0.73), ("delay_ok", 0.09), ("jit", 0.02), ("late", 0.16)])
            peak12 = max(need[:12])
            full_prob = {"robust": 0.45, "delay_ok": 0.75, "jit": 0.85, "late": 0.5}[policy]
            last_landing = N_WEEKS if rng.random() < full_prob else rng.randint(10, 32)
            first_landing = 2
            if policy == "robust":
                stock = 2.3 * peak12 + sum(need[:3]) * rng.uniform(1.0, 1.6)
            elif policy == "delay_ok":
                stock = sum(need[:3]) * rng.uniform(1.2, 2.0) + typical
            elif policy == "jit":
                stock = sum(need[:2]) * rng.uniform(1.0, 1.3) + 0.3 * typical
            else:
                first_landing = rng.randint(4, 10)
                stock = sum(need[:first_landing]) * rng.uniform(0.3, 1.3)
            if rng.random() < 0.45:                       # one or two lots are waiting at the hub
                raben = lot * rng.choice([1, 1, 2])
                if policy != "robust":                    # ...taken out of what the plant holds
                    raben = min(stock, raben)
                    stock -= raben
            stock, raben = num(stock, fractional), num(raben, fractional)
            edi = plan_orders(rng, need, stock, raben, policy, lot, first_landing, last_landing)
            if last_landing >= N_WEEKS and rng.random() < 0.5:      # schedule lines beyond the stock horizon
                k, step = N_WEEKS, rng.choice([1, 2, 2, 3, 4])
                for _ in range(rng.randint(2, 10)):
                    k += step
                    if k < OO_WEEK_COLUMNS:
                        edi[k] = lot * rng.choice([1, 1, 2])
            if rng.random() < 0.25:
                backlog = lot * rng.choice([1, 1, 2])
            if rng.random() < 0.08:
                wk33 = lot
            has_sa = True

        stock, raben = num(stock, fractional), num(raben, fractional)
        proj = [num(v, fractional) for v in project(stock, raben, need, edi)]
        items.append({
            "pn": pn, "family": fam, "segment": seg, "supplier": sup,
            "projects": projects, "price": price, "need": need, "stock": stock, "raben": raben,
            "edi": edi, "backlog": backlog, "wk33": wk33, "proj": proj, "arche": arche,
            "sa": sa_pool[idx] if has_sa else ("Via PO " if rng.random() < 0.35 else None),
            "fractional": fractional, "lot": lot,
            "supplier_pn": (f"{sup['code']}-{rng.randint(10000, 99999)}-{rng.choice('ABCDEFG')}"
                            if rng.random() < 0.27 else None),
        })
    return items


def build_non_raw_rows(rng):
    """Finished / semi-finished goods: present in the forecast sheet, ignored by ANDON."""
    pns = sorted(rng.sample(range(54_000_001, 54_099_999), N_NON_RAW - 24))
    codes = [f"UX{rng.randint(10000, 99999)}" for _ in range(3)] + \
            [f"ZM{rng.randint(10000, 99999)}" for _ in range(21)]
    rows = []
    for pn in pns + sorted(codes):
        need = [0] * N_WEEKS
        if rng.random() < 0.5:
            need = gen_need(rng, rng.lognormvariate(math.log(300), 1.3),
                            weighted(rng, [("steady", 0.6), ("lumpy", 0.4)]), False)
        stock = int(rng.lognormvariate(math.log(400), 1.4)) if rng.random() < 0.42 else 0
        rows.append({"pn": pn, "need": need, "stock": stock, "raben": 0,
                     "proj": project(stock, 0, need, {}), "sup": "Production"})
    return rows


def status_of(proj):
    """Week number of the first negative projected stock, else 'Over Stock '."""
    for i, v in enumerate(proj):
        if v < 0:
            return week_number(i)
    return "Over Stock "


# ----------------------------------------------------------------------------
# Sheet builders (each returns a list of rows; cells are plain values or
# (value, number_format) tuples)
# ----------------------------------------------------------------------------
def need_sheet(raw, non_raw):
    dates = [FIRST_NEED_MONDAY + dt.timedelta(weeks=j) for j in range(N_WEEKS)]
    rows = [[] for _ in range(NEED_HEADER_ROW - 2)]
    rows.append(["Need "] + [(d, "d-mmm") for d in dates])
    header = ["PN"] + [f"CW  {week_number(j + 1):02d}" for j in range(N_WEEKS)]
    header += ["Need Of 14 Weeks ", "All Need ", "All Stock ", "Raben "]
    header += [FIRST_STOCK_WEEK + i for i in range(N_WEEKS)]
    header += ["Status ", "Av Need ", "Target Need ", "PN", "Sup"]
    rows.append(header)

    def line(pn, need, stock, raben, proj, sup, fractional=False):
        n14 = num(sum(need[:14]), fractional)
        av = int(round(sum(need[:14]) / 14))
        return ([pn] + need + [n14, num(sum(need), fractional), stock, raben] + proj
                + [status_of(proj), av, 4 * av, pn, sup])

    body = [line(it["pn"], it["need"], it["stock"], it["raben"], it["proj"],
                 it["supplier"]["name"], it["fractional"]) for it in raw]
    body += [line(r["pn"], r["need"], r["stock"], r["raben"], r["proj"], r["sup"]) for r in non_raw]
    numeric = sorted([b for b in body if isinstance(b[0], int)], key=lambda b: b[0])
    codes = sorted([b for b in body if not isinstance(b[0], int)], key=lambda b: b[0])
    return rows + numeric + codes


def open_orders_sheet(rng, raw, extra_master):
    dates = [FIRST_NEED_MONDAY + dt.timedelta(weeks=k - 1) for k in range(OO_WEEK_COLUMNS)]
    rows = [[] for _ in range(OO_HEADER_ROW - 2)]
    rows.append([None] * 5 + [(d, "d-mmm") for d in dates])
    rows.append(["SA", "Material", "Supplier", "Lead team", "Backlog"]
                + [week_number(k) for k in range(OO_WEEK_COLUMNS)])
    body = []
    for it in raw:
        if not (it["edi"] or it["backlog"] or it["wk33"]):
            continue
        weeks = [None] * OO_WEEK_COLUMNS
        if it["wk33"]:
            weeks[0] = it["wk33"]
        for k, q in it["edi"].items():
            weeks[k] = q
        body.append([it["sa"], it["pn"], it["supplier"]["name"], it["supplier"]["hub_days"],
                     it["backlog"] or None] + weeks)
    for m in extra_master[:12]:                 # a few open agreements on parts without forecast
        weeks = [None] * OO_WEEK_COLUMNS
        weeks[rng.randint(2, 20)] = nice_lot(rng.lognormvariate(math.log(2000), 1.0))
        body.append([m["sa"], m["pn"], m["supplier"]["name"], m["supplier"]["hub_days"], None] + weeks)
    rng.shuffle(body)
    return rows + body


def master_sheet(rng, raw, extra_master, z_rows):
    header = ["PN", "Status ", "SA", "Active / No", "Appro ", "Supplier PN", "Cat ", "supplier",
              "Project ", "Supplier Location ", "All Stock ", "Transit ", "Need 14 Weeks ", "All Need ",
              "Cov till Week", "COV", "Qty Released ", "New Status ", "Bulk ", "mm60", "110+ 140 Value",
              "WIP Value", "Target Per Week", "Average Need", "Target Value", "Inc", "Analyze ABC ",
              "Lead Time ", "MOQ ", "Lot Size ", "Incoterm ", "ETA ", "Contact Person ", "Nbr ",
              "Comment ", "bom"]
    body = []

    def line(pn, seg, sa, active, sup, supplier_pn, projects, stock, transit, need, proj, price, edi_total):
        status, cat, _ = SEGMENTS[seg]
        if active != "Active" and rng.random() < 0.3:
            status = cat = "Obsolete"
        elif rng.random() < 0.08:
            status = "0"
        n14, total = round(sum(need[:14]), 2), round(sum(need), 2)
        avg = round(sum(need[:14]) / 14, 2)
        st = status_of(proj)
        cov_week = "No Need in MD61" if total == 0 else (st.strip() if isinstance(st, str) else st)
        covered = isinstance(st, str)
        per_week = 4 if rng.random() < 0.95 else 2
        described = rng.random() < 0.85
        return [pn, status, sa, active, sup["owner"], supplier_pn, cat, sup["name"], projects,
                sup["location"] if described else None,
                stock, transit, n14, total, cov_week, "Yes" if covered else "NO", edi_total,
                "Yes" if active == "Active" and rng.random() < 0.65 else "NO",
                (sup["tag"] + " ") if rng.random() < 0.06 else None,
                (price, EUR_1), (round(price * stock, 2), EUR_1),
                (round(price * stock * rng.uniform(0.02, 0.2), 2) if rng.random() < 0.1 else 0, EUR_1),
                per_week, avg, (round(avg * per_week * price, 2), EUR_2),
                None, None, sup["lead_time"] if described else None, None, None,
                sup["incoterm"] if described else None, None,
                sup["contact"] if described else None, 1, None, pn]

    in_master = list(raw)
    for it in rng.sample(raw, N_NOT_IN_MASTER):
        in_master.remove(it)
    for it in in_master:
        active = "Active" if sum(it["need"]) > 0 and rng.random() < 0.93 else "No Active "
        body.append(line(it["pn"], it["segment"], it["sa"], active, it["supplier"], it["supplier_pn"],
                         it["projects"], it["stock"], it["raben"], it["need"], it["proj"], it["price"],
                         sum(it["edi"].values())))
    for m in extra_master:
        body.append(line(m["pn"], m["segment"], m["sa"], "No Active ", m["supplier"], None, m["projects"],
                         m["stock"], 0, [0] * N_WEEKS, [m["stock"]] * N_WEEKS, m["price"], 0))
    for z in z_rows:
        body.append(line(z["pn"], z["segment"], "Via PO ", "No Active ", z["supplier"], None, z["projects"],
                         0, 0, [0] * N_WEEKS, [0] * N_WEEKS, 0, 0))
    rng.shuffle(body)

    # ---- summary blocks above the table (same positions as the weekly export) ----
    H = {name.strip(): i for i, name in enumerate(header)}
    owners = [o for o, _ in OWNERS]
    cats = ["HeatPump", "Auto", "Industrial", "New Project", "Obsolete", "Spare Parts"]
    top = [[None] * 36 for _ in range(MASTER_HEADER_ROW - 1)]

    def put(r, c, v):
        top[r][c] = v

    put(0, 12, (dt.datetime(2026, 8, 18, 16, 5), "m/d/yy h:mm"))
    put(1, 12, FIRST_NEED_WEEK)
    put(9, 12, "Appro ")
    for j, c in enumerate(cats):
        put(9, 13 + j, c)
    put(9, 13 + len(cats), "Check ")
    put(10, 1, "Appro "); put(10, 2, "All Parts "); put(10, 4, "Cov "); put(10, 5, "All Parts ")
    covered = sum(1 for b in body if b[H["COV"]] == "Yes")
    put(11, 4, "To be Released "); put(11, 5, len(body) - covered)
    put(12, 4, "Over Stock "); put(12, 5, covered)
    for i, o in enumerate(owners):
        mine = [b for b in body if b[H["Appro"]] == o]
        put(10 + i, 12, o)
        counted = 0
        for j, c in enumerate(cats):
            n = sum(1 for b in mine if b[H["Status"]] == c)
            counted += n
            put(10 + i, 13 + j, n)
        put(10 + i, 13 + len(cats), 0)
        put(11 + i, 1, o); put(11 + i, 2, len(mine))
        put(24 + i, 1, o)
        put(24 + i, 2, sum(1 for b in mine if b[H["Active / No"]] == "Active"))
        put(24 + i, 3, sum(1 for b in mine if b[H["Active / No"]] != "Active"))
        put(24 + i, 4, sum(1 for b in mine if b[H["COV"]] == "Yes"))
    put(23, 1, "Appro "); put(23, 2, "Active"); put(23, 3, "No Active "); put(23, 4, "Over Stock ")
    inventory = round(sum(b[H["110+ 140 Value"]][0] for b in body), 2)
    wip = round(sum(b[H["WIP Value"]][0] for b in body), 2)
    target = round(sum(b[H["Target Value"]][0] for b in body), 2)
    put(22, 7, "Inventory Value "); put(22, 8, (inventory, EUR_2))
    put(23, 7, "WIP Value "); put(23, 8, (wip, EUR_2))
    put(24, 7, "Target Value"); put(24, 8, (target, EUR_2))
    put(26, 7, "Value to explain "); put(26, 8, (round(inventory - target, 2), EUR_2))
    put(35, 20, (inventory, EUR_2))
    return top + [header] + body


def build_master_extras(rng, suppliers):
    used = set()
    def fresh_pn():
        while True:
            pn = rng.randint(14_100_001, 14_199_999)
            if pn not in used:
                used.add(pn)
                return pn
    extra = []
    for _ in range(N_NO_FORECAST):
        fam = weighted(rng, [(f, v[0]) for f, v in FAMILIES.items()])
        _, p_med, p_sig, _, _, spec = FAMILIES[fam]
        seg = weighted(rng, SEGMENT_WEIGHTS)
        tags = SEGMENTS[seg][2]
        extra.append({
            "pn": fresh_pn(), "segment": seg, "supplier": pick_supplier(rng, suppliers, spec),
            "projects": " / ".join(rng.sample(tags, min(len(tags), rng.choice([1, 1, 2])))),
            "stock": int(rng.lognormvariate(math.log(300), 1.3)) if rng.random() < 0.3 else 0,
            "price": round(rng.lognormvariate(math.log(p_med), p_sig), 4) if rng.random() < 0.7 else 0,
            "sa": rng.randint(5_600_100_001, 5_600_199_999) if rng.random() < 0.6 else None,
        })
    z_rows = [{"pn": f"ZM{rng.randint(10000, 99999)}", "segment": "Auto",
               "supplier": pick_supplier(rng, suppliers, "pack"), "projects": "Seat Comfort"}
              for _ in range(N_Z_CODES)]
    return extra, z_rows


# ----------------------------------------------------------------------------
# Writers
# ----------------------------------------------------------------------------
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")


def write_sheet(ws, rows, header_row, widths=None):
    for r, row in enumerate(rows, start=1):
        for c, cell in enumerate(row, start=1):
            if cell is None:
                continue
            value, fmt = cell if isinstance(cell, tuple) else (cell, None)
            x = ws.cell(row=r, column=c, value=value)
            if fmt:
                x.number_format = fmt
            if r == header_row:
                x.font, x.fill = HEADER_FONT, HEADER_FILL
                x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=2)
    for col, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(col)].width = w


def csv_text(cell):
    """Text a spreadsheet would export for a cell (what ANDON's CSV path reads)."""
    if cell is None:
        return ""
    value, fmt = cell if isinstance(cell, tuple) else (cell, None)
    if fmt == "d-mmm":
        return f"{value.day}-{value:%b}"
    if fmt == "m/d/yy h:mm":
        return f"{value.month}/{value.day}/{value:%y} {value.hour}:{value:%M}"
    if fmt == EUR_1:
        return " € -  " if value == 0 else f" € {value:,.1f} "
    if fmt == EUR_2:
        return " € -  " if value == 0 else f" € {value:,.2f} "
    return str(value)


def write_csv(path, rows):
    width = max(len(r) for r in rows)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        for row in rows:
            w.writerow([csv_text(c) for c in row] + [""] * (width - len(row)))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--seed", type=int, default=DEFAULT_SEED)
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent / "sample-data")
    args = ap.parse_args()
    rng = random.Random(args.seed)

    suppliers = build_suppliers(rng)
    raw = build_raw_items(rng, suppliers)
    non_raw = build_non_raw_rows(rng)
    extra_master, z_rows = build_master_extras(rng, suppliers)
    assign_owners(suppliers, raw + extra_master)

    need_rows = need_sheet(raw, non_raw)
    oo_rows = open_orders_sheet(rng, raw, extra_master)
    master_rows = master_sheet(rng, raw, extra_master, z_rows)

    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "csv").mkdir(exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Need"
    write_sheet(ws, need_rows, NEED_HEADER_ROW, {1: 12, 82: 34})
    write_sheet(wb.create_sheet("Open Orders"), oo_rows, OO_HEADER_ROW, {1: 13, 2: 11, 3: 34})
    write_sheet(wb.create_sheet("New R"), master_rows, MASTER_HEADER_ROW,
                {1: 11, 2: 13, 3: 13, 4: 12, 6: 16, 8: 34, 9: 30, 10: 18, 20: 11, 21: 15, 25: 15, 33: 32})
    wb.properties.creator = "ANDON sample-data generator"
    wb.properties.title = "Synthetic supply-planning workbook (fictional data)"
    wb.properties.created = wb.properties.modified = dt.datetime(2026, 8, 18, 16, 5)   # reproducible output
    xlsx = args.out / f"supply_management_{FIRST_NEED_WEEK}-2026_sample.xlsx"
    wb.save(xlsx)

    write_csv(args.out / "csv" / "sample - Need.csv", need_rows)
    write_csv(args.out / "csv" / "sample - Open Orders.csv", oo_rows)
    write_csv(args.out / "csv" / "sample - New R.csv", master_rows)

    with_orders = sum(1 for it in raw if it["edi"])
    print(f"seed {args.seed}: {len(raw)} raw materials ({with_orders} with delivery schedules), "
          f"{len(non_raw)} finished/semi-finished rows, {len(suppliers)} suppliers")
    print(f"written: {xlsx}")
    print(f"written: {args.out / 'csv'} (3 CSV files)")


if __name__ == "__main__":
    main()
