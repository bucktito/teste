import pandas as pd
import json
import re
import shutil
from pathlib import Path

BASE = Path(__file__).parent
EXCEL_PATH = BASE / "source_copy.xlsx"

# Make temporary copies so originals can remain open in Excel
for _src, _dst in [("localizacoes.xlsx","loc_copy.xlsx"), ("metragens.xlsx","met_copy.xlsx")]:
    try: shutil.copy2(BASE / _src, BASE / _dst)
    except PermissionError: pass  # use existing copy if locked

LOC_PATH   = BASE / "loc_copy.xlsx"
MET_PATH   = BASE / "met_copy.xlsx"
OUTPUT_PATH = Path(__file__).parent / "dashboard.html"

xl = pd.ExcelFile(EXCEL_PATH)

def n(v):
    try: return float(v) if pd.notna(v) else 0.0
    except: return 0.0

def categoria(name):
    if "Rua" not in name:
        return "Ed. Barão"
    if "Florêncio" in name or "Flor" in name:
        if "60" in name: return "Rua Florêncio de Abreu, 60"
        return "Rua Florêncio de Abreu, 66/70/74"
    if "25 de Mar" in name:
        if "669" in name: return "Rua 25 de Março, 669"
        if "661" in name: return "Rua 25 de Março, 661"
    return "Rua (outro)"

def unit_type(name):
    nl = name.lower()
    if "loja" in nl: return "Loja"
    if "vitrine" in nl: return "Vitrine"
    if "quiosque" in nl: return "Quiosque"
    if "box" in nl: return "Box"
    if "sala" in nl: return "Sala"
    if "andar" in nl: return "Andar"
    if "rua" in nl: return "Imóvel Rua"
    if re.match(r"\d", name): return "Conjunto"
    return "Outro"

def match_cat(excel_cat):
    """Match a localizacoes.xlsx category name to our CATS keys.
    Uses exact match first so '25 de Março, 661' never matches '25 de Março, 669'.
    """
    ec = excel_cat.strip()
    ec_l = ec.lower()
    # 1. Exact case-insensitive match
    for c in CATS:
        if c.lower() == ec_l:
            return c
    # 2. Barão building — many aliases ("Barão (Ed. Claudina)", etc.)
    if any(k in ec_l for k in ("barão", "barao", "claudina")):
        return "Ed. Barão"
    # 3. Full-string containment (one fully inside the other)
    for c in CATS:
        if c.lower() in ec_l or ec_l in c.lower():
            return c
    return None

CATS = [
    "Rua Florêncio de Abreu, 60",
    "Rua Florêncio de Abreu, 66/70/74",
    "Rua 25 de Março, 669",
    "Rua 25 de Março, 661",
    "Ed. Barão",
]

# ── Sheet 1: Receitas ─────────────────────────────────────────────────────────
raw = pd.read_excel(xl, sheet_name="Receitas", header=None)
total_row_excel = raw[raw[1].astype(str).str.strip() == "TOTAL"].index[0]
data_rows = raw.iloc[7:total_row_excel].copy()

properties = []
for _, row in data_rows.iterrows():
    name = str(row[1]).strip() if pd.notna(row[1]) else ""
    if not name or name == "nan": continue
    aluguel = n(row[2]); encargos = n(row[3]); acordo = n(row[4])
    condominio = n(row[5]); iptu = n(row[6]); honorarios = n(row[10])
    outros = n(row[11]); total = n(row[12])
    cat = categoria(name); utype = unit_type(name)
    pct_admin = round(honorarios / aluguel * 100, 2) if aluguel > 0 else 0.0
    properties.append({
        "name": name, "cat": cat, "type": utype,
        "aluguel": aluguel, "encargos": encargos, "acordo": acordo,
        "condominio": condominio, "iptu": iptu, "honorarios": honorarios,
        "outros": outros, "total": total, "pct_admin": pct_admin,
        "occupied": total > 0, "area": 0.0,
    })

# ── Metragens ─────────────────────────────────────────────────────────────────
met_df = pd.read_excel(MET_PATH)
met_df.columns = ["Categoria", "Unidade", "Tipo", "Area", "Obs"]
met_df["Area"] = pd.to_numeric(met_df["Area"], errors="coerce").fillna(0)
# Build lookup by unit name
area_lookup = {}
for _, row in met_df.iterrows():
    unit = str(row["Unidade"]).strip()
    area = float(row["Area"]) if row["Area"] > 0 else 0.0
    if unit not in area_lookup:
        area_lookup[unit] = area

for p in properties:
    p["area"] = area_lookup.get(p["name"], 0.0)

total_revenue = sum(p["total"] for p in properties)

# ── Aggregate by category ─────────────────────────────────────────────────────
cat_data = {c: {"aluguel": 0, "total": 0, "honorarios": 0, "condominio": 0,
                "iptu": 0, "encargos": 0, "acordo": 0, "outros": 0,
                "units": 0, "occupied": 0, "area": 0, "occupied_area": 0} for c in CATS}
for p in properties:
    c = p["cat"]
    if c not in cat_data: continue
    for k in ["aluguel","total","honorarios","condominio","iptu","encargos","acordo","outros","area"]:
        cat_data[c][k] += p[k]
    cat_data[c]["units"] += 1
    if p["occupied"]:
        cat_data[c]["occupied"] += 1
        cat_data[c]["occupied_area"] += p["area"]

for c in CATS:
    d = cat_data[c]
    d["pct_admin"] = round(d["honorarios"] / d["aluguel"] * 100, 2) if d["aluguel"] > 0 else 0
    d["vacant_area"] = d["area"] - d["occupied_area"]
    occ_a = d["occupied_area"]
    d["aluguel_m2"] = round(d["aluguel"] / occ_a, 2) if occ_a > 0 else 0
    d["occ_area_pct"] = round(d["occupied_area"] / d["area"] * 100, 1) if d["area"] > 0 else 0

# ── Metragem aggregation by type ──────────────────────────────────────────────
type_met = {}
for p in properties:
    t = p["type"]
    if t not in type_met:
        type_met[t] = {"aluguel": 0, "area": 0, "units": 0, "occupied": 0,
                       "occupied_area": 0, "vacant_area": 0}
    type_met[t]["units"] += 1
    type_met[t]["area"] += p["area"]
    if p["occupied"]:
        type_met[t]["aluguel"] += p["aluguel"]
        type_met[t]["occupied"] += 1
        type_met[t]["occupied_area"] += p["area"]
    else:
        type_met[t]["vacant_area"] += p["area"]

for t, d in type_met.items():
    d["aluguel_m2"] = round(d["aluguel"] / d["occupied_area"], 2) if d["occupied_area"] > 0 else 0
    d["vacant_units"] = d["units"] - d["occupied"]
    d["occ_pct"] = round(d["occupied"] / d["units"] * 100, 1) if d["units"] > 0 else 0
    d["occ_area_pct"] = round(d["occupied_area"] / d["area"] * 100, 1) if d["area"] > 0 else 0

# ── Metragem aggregation by (cat, type) for Estimativa ───────────────────────
cat_type_stats = {}
for p in properties:
    key = f"{p['cat']}||{p['type']}"
    if key not in cat_type_stats:
        cat_type_stats[key] = {"cat": p["cat"], "type": p["type"],
                               "aluguel": 0, "area": 0, "occupied_area": 0,
                               "vacant_area": 0, "units": 0, "occupied": 0}
    d = cat_type_stats[key]
    d["units"] += 1; d["area"] += p["area"]
    if p["occupied"]:
        d["aluguel"] += p["aluguel"]; d["occupied"] += 1; d["occupied_area"] += p["area"]
    else:
        d["vacant_area"] += p["area"]

for key, d in cat_type_stats.items():
    d["aluguel_m2"] = round(d["aluguel"] / d["occupied_area"], 2) if d["occupied_area"] > 0 else 0
    d["vacant"] = d["units"] - d["occupied"]
    d["occ_pct"] = round(d["occupied"] / d["units"] * 100, 1) if d["units"] > 0 else 0

# ── Portfolio consolidated KPIs ───────────────────────────────────────────────
total_aluguel    = sum(p["aluguel"] for p in properties)
total_honorarios = sum(p["honorarios"] for p in properties)
total_units      = sum(d["units"] for d in cat_data.values())
total_occupied   = sum(d["occupied"] for d in cat_data.values())
total_vacant     = total_units - total_occupied
total_area       = sum(p["area"] for p in properties)
total_occ_area   = sum(p["area"] for p in properties if p["occupied"])
total_vac_area   = total_area - total_occ_area
occ_pct          = round(total_occupied / total_units * 100, 1) if total_units else 0
vac_pct          = round(100 - occ_pct, 1)
pct_hon_total    = round(total_honorarios / total_aluguel * 100, 2) if total_aluguel else 0
aluguel_m2_port  = round(total_aluguel / total_occ_area, 2) if total_occ_area else 0
occ_area_pct     = round(total_occ_area / total_area * 100, 1) if total_area else 0

# ── Revenue components ────────────────────────────────────────────────────────
comp_totals = {}
for key in ["aluguel", "encargos", "acordo", "condominio", "iptu", "honorarios", "outros"]:
    v = sum(p[key] for p in properties)
    if v > 0: comp_totals[key.capitalize()] = round(v, 2)

# ── Despesas ──────────────────────────────────────────────────────────────────
raw_desp = pd.read_excel(xl, sheet_name="Dist. Despesas", header=None)
total_desp_row = raw_desp[raw_desp[1].astype(str).str.contains("TOTAL DE DESPESAS", na=False)]
total_expenses = abs(float(total_desp_row.iloc[0][2])) if not total_desp_row.empty else 0.0

# ── Saldo Individual ──────────────────────────────────────────────────────────
raw_saldo = pd.read_excel(xl, sheet_name="Saldo Individual", header=None)
balances = []
for _, row in raw_saldo.iterrows():
    name = str(row[1]).strip() if pd.notna(row[1]) else ""
    if not name or name == "nan" or any(s in name.upper() for s in {"SALDO", "NAN"}): continue
    try:
        balances.append({"name": name, "prev": n(row[2]), "receitas": n(row[3]),
                         "despesas": n(row[4]), "atual": n(row[5])})
    except: continue

# ── Vacancy by type ───────────────────────────────────────────────────────────
barao_props = [p for p in properties if p["cat"] == "Ed. Barão"]
type_stats_vac = {}
for p in barao_props:
    t = p["type"]
    if t not in type_stats_vac: type_stats_vac[t] = {"total": 0, "occupied": 0}
    type_stats_vac[t]["total"] += 1
    if p["occupied"]: type_stats_vac[t]["occupied"] += 1

vacancy_by_type = [
    {"type": t, "total": v["total"], "occupied": v["occupied"],
     "vacant": v["total"] - v["occupied"],
     "pct": round(v["occupied"] / v["total"] * 100, 1) if v["total"] > 0 else 0}
    for t, v in type_stats_vac.items()
]

# ── Admin units ───────────────────────────────────────────────────────────────
admin_units = sorted(
    [p for p in properties if p["aluguel"] > 0 and p["honorarios"] > 0],
    key=lambda x: x["aluguel"], reverse=True
)[:30]

# ── Locations — flexible matching ─────────────────────────────────────────────
loc_df = pd.read_excel(LOC_PATH)
locations = []
for _, row in loc_df.iterrows():
    raw_cat = str(row.iloc[0]).strip()
    addr    = str(row.iloc[1]).strip()
    lat     = float(row.iloc[2]) if pd.notna(row.iloc[2]) else None
    lng     = float(row.iloc[3]) if pd.notna(row.iloc[3]) else None
    if not lat or not lng: continue
    cat = match_cat(raw_cat)
    if cat and cat in cat_data:
        locations.append({
            "cat": cat, "address": addr, "lat": lat, "lng": lng,
            "aluguel":  round(cat_data[cat]["aluguel"], 2),
            "total":    round(cat_data[cat]["total"], 2),
            "units":    cat_data[cat]["units"],
            "occupied": cat_data[cat]["occupied"],
        })

barao_total_units    = cat_data["Ed. Barão"]["units"]
barao_occupied_units = cat_data["Ed. Barão"]["occupied"]
barao_vacant_units   = barao_total_units - barao_occupied_units
barao_occ_pct        = round(barao_occupied_units / barao_total_units * 100, 1) if barao_total_units else 0

# ── Units data for JS (full, for filter + estimativa) ────────────────────────
units_js = [{"name": p["name"], "cat": p["cat"], "type": p["type"],
             "ct_key": f"{p['cat']}||{p['type']}",
             "aluguel": round(p["aluguel"],2), "total": round(p["total"],2),
             "honorarios": round(p["honorarios"],2), "condominio": round(p["condominio"],2),
             "iptu": round(p["iptu"],2), "encargos": round(p["encargos"],2),
             "acordo": round(p["acordo"],2), "outros": round(p["outros"],2),
             "area": p["area"], "occupied": p["occupied"]} for p in properties]

# ── cat_type_stats for JS ─────────────────────────────────────────────────────
cat_type_js = {k: {ki: round(v,2) if isinstance(v, float) else v
                   for ki, v in d.items()}
               for k, d in cat_type_stats.items()}

# ── type_met for JS ───────────────────────────────────────────────────────────
type_met_js = {t: {k: round(v,2) if isinstance(v, float) else v for k,v in d.items()}
               for t,d in type_met.items()}

data = {
    "month": "Maio/2026",
    "total_revenue": round(total_revenue, 2),
    "total_expenses": round(total_expenses, 2),
    "net_balance": round(total_revenue - total_expenses, 2),
    "total_aluguel": round(total_aluguel, 2),
    "total_honorarios": round(total_honorarios, 2),
    "pct_hon_total": pct_hon_total,
    "total_units": total_units,
    "total_occupied": total_occupied,
    "total_vacant": total_vacant,
    "total_area": total_area,
    "total_occ_area": round(total_occ_area, 1),
    "total_vac_area": round(total_vac_area, 1),
    "occ_pct": occ_pct,
    "vac_pct": vac_pct,
    "occ_area_pct": occ_area_pct,
    "aluguel_m2_port": aluguel_m2_port,
    "cat_data": {c: {k: round(v, 2) if isinstance(v, float) else v for k, v in d.items()} for c, d in cat_data.items()},
    "comp_totals": comp_totals,
    "balances": balances,
    "properties": units_js,
    "admin_units": [{"name": p["name"], "cat": p["cat"], "aluguel": round(p["aluguel"],2),
                     "honorarios": round(p["honorarios"],2), "pct_admin": p["pct_admin"]} for p in admin_units],
    "locations": locations,
    "vacancy_by_type": vacancy_by_type,
    "barao_total_units": barao_total_units,
    "barao_occupied_units": barao_occupied_units,
    "barao_vacant_units": barao_vacant_units,
    "barao_occ_pct": barao_occ_pct,
    "type_met": type_met_js,
    "cat_type": cat_type_js,
    "CATS": CATS,
}

CAT_COLORS = {
    "Rua Florêncio de Abreu, 60":       "#4f8ef7",
    "Rua Florêncio de Abreu, 66/70/74": "#5dca8a",
    "Rua 25 de Março, 669":             "#f7c948",
    "Rua 25 de Março, 661":             "#fb923c",
    "Ed. Barão":                        "#a78bfa",
}
PILL_COLORS_LIST = ["#4f8ef7","#5dca8a","#f7c948","#fb923c","#a78bfa"]

# ── Static HTML rows ──────────────────────────────────────────────────────────
cat_rows = ""
for cat in CATS:
    d = cat_data[cat]; color = CAT_COLORS[cat]
    occ_pct_c = round(d["occupied"] / d["units"] * 100, 1) if d["units"] else 0
    cat_rows += f"""
    <tr>
      <td><span class="tag" style="background:{color}22;color:{color}">{cat}</span></td>
      <td>R$ {d['aluguel']:,.2f}</td><td>R$ {d['total']:,.2f}</td>
      <td>R$ {d['honorarios']:,.2f}</td><td>{d['pct_admin']:.2f}%</td>
      <td>{d['occupied']}/{d['units']} ({occ_pct_c:.0f}%)</td>
    </tr>"""

bal_rows = ""
for b in balances:
    pn = "positive" if b["atual"] >= 0 else "negative"
    bal_rows += f"""
    <tr>
      <td>{b['name']}</td>
      <td class="{'positive' if b['prev']>=0 else 'negative'}">R$ {b['prev']:,.2f}</td>
      <td class="positive">R$ {b['receitas']:,.2f}</td>
      <td class="negative">R$ {b['despesas']:,.2f}</td>
      <td class="{pn}"><strong>R$ {b['atual']:,.2f}</strong></td>
    </tr>"""

admin_rows = ""
for p in data["admin_units"]:
    color = CAT_COLORS.get(p["cat"], "#888")
    admin_rows += f"""
    <tr>
      <td>{p['name']}</td>
      <td><span class="tag" style="background:{color}22;color:{color}">{p['cat']}</span></td>
      <td>R$ {p['aluguel']:,.2f}</td><td>R$ {p['honorarios']:,.2f}</td>
      <td><div class="pct-label">{p['pct_admin']:.2f}%</div>
          <div class="bar"><div class="bar-fill" style="width:{min(p['pct_admin']/10*100,100):.0f}%;background:{'#5dca8a' if p['pct_admin']<=6 else '#f76c6c'}"></div></div></td>
    </tr>"""

vac_rows = ""
for v in sorted(data["vacancy_by_type"], key=lambda x: x["total"], reverse=True):
    bar = v["pct"]; color = "#5dca8a" if bar >= 80 else "#f7c948" if bar >= 60 else "#f76c6c"
    vac_rows += f"""
    <tr>
      <td>{v['type']}</td><td>{v['total']}</td>
      <td class="positive">{v['occupied']}</td><td class="negative">{v['vacant']}</td>
      <td><div class="pct-label" style="color:{color}">{v['pct']:.1f}%</div>
          <div class="bar"><div class="bar-fill" style="width:{bar:.1f}%;background:{color}"></div></div></td>
    </tr>"""

# Metragem category rows
met_cat_rows = ""
for cat in CATS:
    d = cat_data[cat]; color = CAT_COLORS[cat]
    met_cat_rows += f"""
    <tr>
      <td><span class="tag" style="background:{color}22;color:{color}">{cat}</span></td>
      <td>{d['area']:,.0f} m²</td>
      <td>{d['occupied_area']:,.0f} m²</td>
      <td>{d['vacant_area']:,.0f} m²</td>
      <td>{'R$ '+f"{d['aluguel_m2']:,.2f}"+'/m²' if d['aluguel_m2'] else '—'}</td>
      <td>{d['occ_area_pct']:.1f}%</td>
    </tr>"""

# Metragem type rows
met_type_rows = ""
TYPE_COLORS = ["#4f8ef7","#5dca8a","#f7c948","#fb923c","#a78bfa","#38bdf8","#f472b6","#34d399"]
for i, (t, d) in enumerate(sorted(type_met.items(), key=lambda x: x[1]["area"], reverse=True)):
    color = TYPE_COLORS[i % len(TYPE_COLORS)]
    met_type_rows += f"""
    <tr>
      <td><span class="tag" style="background:{color}22;color:{color}">{t}</span></td>
      <td>{d['area']:,.0f} m²</td>
      <td>{d['occupied_area']:,.0f} m²</td>
      <td>{d['vacant_area']:,.0f} m²</td>
      <td>{'R$ '+f"{d['aluguel_m2']:,.2f}"+'/m²' if d['aluguel_m2'] else '—'}</td>
      <td>{d['occ_area_pct']:.1f}%</td>
      <td>{d['occupied']}/{d['units']}</td>
    </tr>"""

# Estimativa type input rows (JS-rendered, just list types)
est_types = sorted(type_met.keys(), key=lambda t: type_met[t]["area"], reverse=True)

map_markers_js = "[\n" + ",\n".join(
    f'  {{cat:"{m["cat"]}",lat:{m["lat"]},lng:{m["lng"]},addr:"{m["address"]}",aluguel:{m["aluguel"]},total:{m["total"]},units:{m["units"]},occupied:{m["occupied"]}}}'
    for m in data["locations"]
) + "\n]"

cat_colors_js  = json.dumps(CAT_COLORS)
type_colors_js = json.dumps({t: TYPE_COLORS[i % len(TYPE_COLORS)] for i, t in enumerate(sorted(type_met.keys(), key=lambda x: type_met[x]["area"], reverse=True))})

# ── HTML ──────────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Dashboard – Fam. Keutenedjian | {data['month']}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
:root{{
  --bg:#0d0f18;--surface:#141720;--surface2:#1c1f2e;--border:#252838;
  --accent:#4f8ef7;--green:#5dca8a;--yellow:#f7c948;--red:#f76c6c;--purple:#a78bfa;--orange:#fb923c;
  --text:#e2e8f0;--muted:#8892a4;--muted2:#5a6277;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,Arial,sans-serif;min-height:100vh}}

.header-wrap{{background:linear-gradient(135deg,#0d1526 0%,#141b35 50%,#0f1520 100%);border-bottom:1px solid var(--border);position:relative;overflow:hidden}}
.header-wrap::before{{content:'';position:absolute;inset:0;background:radial-gradient(ellipse 60% 80% at 80% 50%,rgba(79,142,247,.08) 0%,transparent 70%);pointer-events:none}}
.header-top{{display:flex;align-items:center;justify-content:space-between;padding:20px 36px 16px}}
.header-left{{display:flex;align-items:center;gap:16px}}
.header-monogram{{width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,#4f8ef7,#a78bfa);display:flex;align-items:center;justify-content:center;font-size:1.1rem;font-weight:800;color:#fff;box-shadow:0 4px 20px rgba(79,142,247,.35);flex-shrink:0}}
.header-title h1{{font-size:1.15rem;font-weight:700;letter-spacing:-.01em}}
.header-title p{{font-size:.72rem;color:var(--muted);margin-top:2px;letter-spacing:.03em}}
.header-right{{display:flex;align-items:center;gap:12px}}
.badge-month{{background:linear-gradient(135deg,rgba(79,142,247,.2),rgba(167,139,250,.2));border:1px solid rgba(79,142,247,.35);color:#a5c4fd;font-size:.78rem;font-weight:600;padding:5px 14px;border-radius:20px;letter-spacing:.04em}}
.badge-live{{display:flex;align-items:center;gap:6px;font-size:.7rem;color:var(--green);font-weight:600}}
.badge-live .dot{{width:6px;height:6px;border-radius:50%;background:var(--green);animation:pulse 2s infinite}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.3}}}}

.portfolio-strip{{display:grid;grid-template-columns:repeat(6,1fr);border-top:1px solid var(--border)}}
.ps-item{{padding:14px 20px;border-right:1px solid var(--border);position:relative}}
.ps-item:last-child{{border-right:none}}
.ps-item::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--accent-c,var(--accent));opacity:.6}}
.ps-label{{font-size:.63rem;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:5px}}
.ps-value{{font-size:1.18rem;font-weight:700;line-height:1;color:var(--ps-color,var(--text))}}
.ps-sub{{font-size:.68rem;color:var(--muted2);margin-top:3px}}

.tabs{{background:var(--surface);border-bottom:1px solid var(--border);display:flex;padding:0 36px;overflow-x:auto}}
.tab{{padding:13px 18px;font-size:.8rem;font-weight:600;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;transition:color .2s,border-color .2s;white-space:nowrap}}
.tab:hover{{color:var(--text)}}
.tab.active{{color:var(--accent);border-bottom-color:var(--accent)}}

.page{{display:none;padding:24px 36px}}
.page.active{{display:block}}

.filter-bar{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:18px;align-items:center}}
.filter-bar .fb-label{{font-size:.68rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-right:2px}}
.pill{{padding:5px 13px;border-radius:20px;font-size:.74rem;font-weight:600;cursor:pointer;border:1px solid var(--border);color:var(--muted);background:transparent;transition:all .15s;white-space:nowrap;user-select:none}}
.pill:hover{{border-color:var(--accent);color:var(--text)}}
.pill.active{{color:#fff;border-color:transparent}}

.dkpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:13px;margin-bottom:18px}}
.dkpi{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:15px 18px;position:relative;overflow:hidden}}
.dkpi::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;background:var(--dkpi-color,var(--accent));opacity:.5}}
.dkpi .lbl{{font-size:.63rem;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);margin-bottom:6px}}
.dkpi .val{{font-size:1.28rem;font-weight:700;color:var(--dkpi-color,var(--text))}}
.dkpi .sub{{font-size:.66rem;color:var(--muted2);margin-top:3px}}

.chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:20px}}
.chart-card h2{{font-size:.7rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;margin-bottom:16px}}
.charts-row{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}}
.charts-row-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;margin-bottom:18px}}
.mb{{margin-bottom:18px}}

table{{width:100%;border-collapse:collapse;font-size:.81rem}}
thead tr{{border-bottom:1px solid var(--border)}}
th{{text-align:left;padding:8px 11px;color:var(--muted);font-weight:700;font-size:.64rem;text-transform:uppercase;letter-spacing:.07em}}
td{{padding:8px 11px;border-bottom:1px solid rgba(255,255,255,.035)}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:rgba(255,255,255,.02)}}
.tag{{display:inline-block;font-size:.66rem;padding:2px 8px;border-radius:6px;font-weight:700}}
.bar{{height:5px;border-radius:3px;background:var(--border);margin-top:4px}}
.bar-fill{{height:100%;border-radius:3px}}
.pct-label{{font-size:.74rem;color:var(--muted)}}
.positive{{color:var(--green)}} .negative{{color:var(--red)}}

/* Estimativa */
.est-grid{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}}
.est-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:20px}}
.est-card h3{{font-size:.7rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;margin-bottom:14px}}
.est-row{{display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid rgba(255,255,255,.04)}}
.est-row:last-child{{border-bottom:none}}
.est-label{{flex:1;font-size:.8rem;color:var(--text)}}
.est-type-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.est-input{{width:90px;background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:4px 8px;border-radius:6px;font-size:.8rem;text-align:right}}
.est-input:focus{{outline:none;border-color:var(--accent)}}
.est-unit{{font-size:.7rem;color:var(--muted);width:28px}}
.est-slider{{width:100px;accent-color:var(--accent)}}
.est-pct{{font-size:.78rem;color:var(--accent);min-width:36px;text-align:right}}
.est-result-strip{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}}
.est-res{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:14px 18px}}
.est-res .er-lbl{{font-size:.63rem;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);margin-bottom:6px}}
.est-res .er-val{{font-size:1.2rem;font-weight:700}}

#map{{height:440px;border-radius:10px;border:1px solid var(--border)}}
.leaflet-popup-content-wrapper{{background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:10px;box-shadow:0 8px 32px rgba(0,0,0,.6)}}
.leaflet-popup-tip{{background:var(--surface)}}
.leaflet-popup-content{{margin:13px 16px;font-size:.8rem}}
.leaflet-popup-content strong{{color:var(--accent);font-size:.86rem}}
.pop-row{{display:flex;justify-content:space-between;gap:20px;margin-top:4px}}
.pop-row span:last-child{{color:var(--green);font-weight:600}}

@media(max-width:900px){{
  .portfolio-strip{{grid-template-columns:repeat(3,1fr)}}
  .charts-row,.charts-row-3,.est-grid{{grid-template-columns:1fr}}
  .est-result-strip{{grid-template-columns:repeat(2,1fr)}}
  .page{{padding:14px 16px}}
  .tabs{{padding:0 12px}}
}}
</style>
</head>
<body>

<div class="header-wrap">
  <div class="header-top">
    <div class="header-left">
      <div class="header-monogram">K</div>
      <div class="header-title">
        <h1>Família Keutenedjian — Gestão Imobiliária</h1>
        <p>Prestação de Contas · Carteira de Ativos Imobiliários</p>
      </div>
    </div>
    <div class="header-right">
      <div class="badge-live"><div class="dot"></div>Atualizado</div>
      <div class="badge-month">{data['month']}</div>
    </div>
  </div>
  <div class="portfolio-strip">
    <div class="ps-item" style="--accent-c:var(--red);--ps-color:var(--red)">
      <div class="ps-label">Vacância</div>
      <div class="ps-value">{data['vac_pct']:.1f}%</div>
      <div class="ps-sub">{data['total_vacant']} unidades vagas</div>
    </div>
    <div class="ps-item" style="--accent-c:var(--green);--ps-color:var(--green)">
      <div class="ps-label">Ocupação</div>
      <div class="ps-value">{data['occ_pct']:.1f}%</div>
      <div class="ps-sub">{data['total_occupied']}/{data['total_units']} unidades</div>
    </div>
    <div class="ps-item" style="--accent-c:var(--accent);--ps-color:var(--accent)">
      <div class="ps-label">Aluguel Total</div>
      <div class="ps-value" style="font-size:.92rem">R$ {data['total_aluguel']:,.0f}</div>
      <div class="ps-sub">Contratos ativos</div>
    </div>
    <div class="ps-item" style="--accent-c:var(--yellow);--ps-color:var(--yellow)">
      <div class="ps-label">Receita Bruta</div>
      <div class="ps-value" style="font-size:.92rem">R$ {data['total_revenue']:,.0f}</div>
      <div class="ps-sub">Aluguel + encargos + IPTU</div>
    </div>
    <div class="ps-item" style="--accent-c:var(--orange);--ps-color:var(--orange)">
      <div class="ps-label">Honorários</div>
      <div class="ps-value" style="font-size:.92rem">R$ {data['total_honorarios']:,.0f}</div>
      <div class="ps-sub">{data['pct_hon_total']:.2f}% do aluguel</div>
    </div>
    <div class="ps-item" style="--accent-c:var(--purple);--ps-color:var(--purple)">
      <div class="ps-label">Aluguel / m²</div>
      <div class="ps-value">R$ {data['aluguel_m2_port']:,.2f}</div>
      <div class="ps-sub">{data['total_occ_area']:,.0f} m² ocupados</div>
    </div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" onclick="switchTab('overview')">Visão Geral</div>
  <div class="tab" onclick="switchTab('admin')">Comparativo Administração</div>
  <div class="tab" onclick="switchTab('mapa')">Mapa</div>
  <div class="tab" onclick="switchTab('vacancia')">Vacância e Ocupação</div>
  <div class="tab" onclick="switchTab('metragem')">Metragem</div>
  <div class="tab" onclick="switchTab('estimativa')">Estimativa</div>
  <div class="tab" onclick="switchTab('saldos')">Saldos Individuais</div>
</div>

<!-- ══ VISÃO GERAL ════════════════════════════════════════════════════════════ -->
<div class="page active" id="page-overview">
  <div class="filter-bar">
    <span class="fb-label">Filtrar ativos:</span>
    <div class="pill active" id="pill-all" onclick="toggleAll()">Todos</div>
    {''.join(f'<div class="pill" id="pill-{i}" onclick="toggleCat({i})">{["Fl. Abreu 60","Fl. Abreu 66/70/74","25 de Março 669","25 de Março 661","Ed. Barão"][i]}</div>' for i in range(5))}
  </div>
  <div class="dkpis">
    <div class="dkpi" style="--dkpi-color:var(--accent)"><div class="lbl">Aluguel</div><div class="val" id="dk-aluguel">—</div></div>
    <div class="dkpi" style="--dkpi-color:var(--yellow)"><div class="lbl">Receita Total</div><div class="val" id="dk-receita">—</div></div>
    <div class="dkpi" style="--dkpi-color:var(--green)"><div class="lbl">Honorários</div><div class="val" id="dk-honorarios">—</div></div>
    <div class="dkpi" style="--dkpi-color:var(--purple)"><div class="lbl">Aluguel / m²</div><div class="val" id="dk-m2">—</div></div>
    <div class="dkpi" style="--dkpi-color:var(--green)"><div class="lbl">Unidades Ocupadas</div><div class="val" id="dk-ocupadas">—</div></div>
    <div class="dkpi" style="--dkpi-color:var(--red)"><div class="lbl">Unidades Vagas</div><div class="val" id="dk-vagas">—</div></div>
  </div>
  <div class="charts-row mb">
    <div class="chart-card">
      <h2>Aluguel Total por Categoria de Imóvel</h2>
      <div style="height:250px"><canvas id="catAluguelChart"></canvas></div>
    </div>
    <div class="chart-card">
      <h2>Composição da Receita</h2>
      <div style="height:250px"><canvas id="compChart"></canvas></div>
    </div>
  </div>
  <div class="chart-card mb">
    <h2>Resumo por Categoria</h2>
    <table><thead><tr><th>Categoria</th><th>Aluguel</th><th>Receita Total</th><th>Honorários</th><th>% Admin</th><th>Ocupação</th></tr></thead>
    <tbody>{cat_rows}</tbody></table>
  </div>
  <div class="chart-card">
    <h2 id="h-top-table">Top Unidades por Receita</h2>
    <table><thead><tr><th>Unidade</th><th>Categoria</th><th>Aluguel</th><th>Receita Total</th><th>% do Total</th></tr></thead>
    <tbody id="top-table-body"></tbody></table>
  </div>
</div>

<!-- ══ COMPARATIVO ADMINISTRAÇÃO ═════════════════════════════════════════════ -->
<div class="page" id="page-admin">
  <div class="dkpis" style="grid-template-columns:repeat(5,1fr)">
    {''.join(f'''<div class="dkpi" style="--dkpi-color:{CAT_COLORS[cat]}">
      <div class="lbl">{cat}</div><div class="val">{data['cat_data'][cat]['pct_admin']:.2f}%</div>
      <div class="sub">R$ {data['cat_data'][cat]['honorarios']:,.2f}</div></div>''' for cat in CATS)}
  </div>
  <div class="charts-row mb">
    <div class="chart-card"><h2>% Honorários por Categoria</h2><div style="height:270px"><canvas id="adminPctChart"></canvas></div></div>
    <div class="chart-card"><h2>Honorários vs Aluguel (R$)</h2><div style="height:270px"><canvas id="adminAbsChart"></canvas></div></div>
  </div>
  <div class="chart-card">
    <h2>Detalhamento por Unidade — Top 30 por Aluguel</h2>
    <table><thead><tr><th>Unidade</th><th>Categoria</th><th>Aluguel</th><th>Honorários</th><th>% Admin</th></tr></thead>
    <tbody>{admin_rows}</tbody></table>
  </div>
</div>

<!-- ══ MAPA ═══════════════════════════════════════════════════════════════════ -->
<div class="page" id="page-mapa">
  <div class="dkpis" style="grid-template-columns:repeat(5,1fr);margin-bottom:16px">
    {''.join(f'''<div class="dkpi" style="--dkpi-color:{CAT_COLORS[cat]}">
      <div class="lbl">{cat}</div>
      <div class="val" style="font-size:.88rem">R$ {data['cat_data'][cat]['aluguel']:,.0f}</div>
      <div class="sub">Aluguel mensal</div></div>''' for cat in CATS)}
  </div>
  <div class="chart-card">
    <h2>Localização dos Imóveis — tamanho do círculo = aluguel mensal</h2>
    <div id="map"></div>
    <p style="font-size:.7rem;color:var(--muted2);margin-top:10px">
      ⚠️ Coordenadas lidas de <strong>localizacoes.xlsx</strong> — atualize o arquivo para corrigir posições.
    </p>
  </div>
</div>

<!-- ══ VACÂNCIA E OCUPAÇÃO ════════════════════════════════════════════════════ -->
<div class="page" id="page-vacancia">
  <div class="dkpis">
    <div class="dkpi" style="--dkpi-color:var(--yellow)"><div class="lbl">Total Unidades Barão</div><div class="val">{data['barao_total_units']}</div></div>
    <div class="dkpi" style="--dkpi-color:var(--green)"><div class="lbl">Unidades Ocupadas</div><div class="val">{data['barao_occupied_units']}</div></div>
    <div class="dkpi" style="--dkpi-color:var(--red)"><div class="lbl">Unidades Vagas</div><div class="val">{data['barao_vacant_units']}</div></div>
    <div class="dkpi" style="--dkpi-color:{'var(--green)' if data['barao_occ_pct']>=80 else 'var(--yellow)'}"><div class="lbl">Taxa de Ocupação</div><div class="val">{data['barao_occ_pct']:.1f}%</div></div>
    <div class="dkpi" style="--dkpi-color:var(--red)"><div class="lbl">Taxa de Vacância</div><div class="val">{100-data['barao_occ_pct']:.1f}%</div></div>
  </div>
  <div class="charts-row mb">
    <div class="chart-card"><h2>Ocupação Ed. Barão por Tipo</h2><div style="height:290px"><canvas id="vaccTypeChart"></canvas></div></div>
    <div class="chart-card"><h2>Ocupação × Vacância — Ed. Barão</h2><div style="height:290px"><canvas id="vaccDonut"></canvas></div></div>
  </div>
  <div class="charts-row mb">
    <div class="chart-card">
      <h2>Imóveis Rua — Status</h2>
      <table><thead><tr><th>Imóvel</th><th>Tipo</th><th>Aluguel</th><th>Status</th></tr></thead>
      <tbody>{''.join(f"""<tr><td>{cat}</td><td>Contrato Único</td><td>R$ {data['cat_data'][cat]['aluguel']:,.2f}</td>
        <td><span class="tag" style="background:#5dca8a22;color:#5dca8a">Ocupado</span></td></tr>""" for cat in CATS if cat!="Ed. Barão")}</tbody></table>
    </div>
    <div class="chart-card">
      <h2>Distribuição de Vacância por Tipo</h2>
      <table><thead><tr><th>Tipo</th><th>Total</th><th>Ocupadas</th><th>Vagas</th><th>Ocupação</th></tr></thead>
      <tbody>{vac_rows}</tbody></table>
    </div>
  </div>
</div>

<!-- ══ METRAGEM ═══════════════════════════════════════════════════════════════ -->
<div class="page" id="page-metragem">
  <div class="dkpis">
    <div class="dkpi" style="--dkpi-color:var(--accent)"><div class="lbl">Área Total Carteira</div><div class="val">{data['total_area']:,.0f} m²</div></div>
    <div class="dkpi" style="--dkpi-color:var(--green)"><div class="lbl">Área Ocupada</div><div class="val">{data['total_occ_area']:,.0f} m²</div><div class="sub">{data['occ_area_pct']:.1f}% do total</div></div>
    <div class="dkpi" style="--dkpi-color:var(--red)"><div class="lbl">Área Vaga</div><div class="val">{data['total_vac_area']:,.0f} m²</div><div class="sub">{100-data['occ_area_pct']:.1f}% do total</div></div>
    <div class="dkpi" style="--dkpi-color:var(--purple)"><div class="lbl">Aluguel / m² (carteira)</div><div class="val">R$ {data['aluguel_m2_port']:,.2f}</div><div class="sub">Média ponderada</div></div>
  </div>
  <div class="charts-row mb">
    <div class="chart-card"><h2>Aluguel / m² por Categoria</h2><div style="height:260px"><canvas id="metCatM2Chart"></canvas></div></div>
    <div class="chart-card"><h2>Representatividade por Metragem</h2><div style="height:260px"><canvas id="metAreaDonut"></canvas></div></div>
  </div>
  <div class="charts-row mb">
    <div class="chart-card"><h2>Ocupação / Vacância por m² (por Categoria)</h2><div style="height:260px"><canvas id="metOccAreaChart"></canvas></div></div>
    <div class="chart-card"><h2>Aluguel / m² por Tipo de Imóvel</h2><div style="height:260px"><canvas id="metTypeM2Chart"></canvas></div></div>
  </div>
  <div class="charts-row mb">
    <div class="chart-card">
      <h2>Área e Aluguel/m² por Categoria</h2>
      <table><thead><tr><th>Categoria</th><th>Área Total</th><th>Área Ocupada</th><th>Área Vaga</th><th>Aluguel/m²</th><th>Ocupação Área</th></tr></thead>
      <tbody>{met_cat_rows}</tbody></table>
    </div>
    <div class="chart-card">
      <h2>Área e Aluguel/m² por Tipo</h2>
      <table><thead><tr><th>Tipo</th><th>Área Total</th><th>Área Ocupada</th><th>Área Vaga</th><th>Aluguel/m²</th><th>Ocup. Área</th><th>Unidades</th></tr></thead>
      <tbody>{met_type_rows}</tbody></table>
    </div>
  </div>
</div>

<!-- ══ ESTIMATIVA ════════════════════════════════════════════════════════════ -->
<div class="page" id="page-estimativa">
  <div class="est-result-strip" id="est-result-strip">
    <div class="est-res"><div class="er-lbl">Aluguel Estimado</div><div class="er-val" id="er-aluguel" style="color:var(--accent)">—</div></div>
    <div class="est-res"><div class="er-lbl">vs. Atual</div><div class="er-val" id="er-diff">—</div></div>
    <div class="est-res"><div class="er-lbl">Área Ocupada Estimada</div><div class="er-val" id="er-area" style="color:var(--purple)">—</div></div>
    <div class="est-res"><div class="er-lbl">Unidades Estimadas</div><div class="er-val" id="er-units" style="color:var(--green)">—</div></div>
  </div>
  <div class="est-grid">
    <div class="est-card">
      <h3>Preço por m² por Tipo de Imóvel</h3>
      <div id="est-price-inputs"></div>
    </div>
    <div class="est-card">
      <h3>Cenário de Ocupação por Categoria</h3>
      <div id="est-occ-inputs"></div>
    </div>
  </div>
  <div class="charts-row">
    <div class="chart-card"><h2>Aluguel Estimado por Categoria</h2><div style="height:260px"><canvas id="estCatChart"></canvas></div></div>
    <div class="chart-card"><h2>Estimado vs Atual por Categoria</h2><div style="height:260px"><canvas id="estDiffChart"></canvas></div></div>
  </div>
</div>

<!-- ══ SALDOS ════════════════════════════════════════════════════════════════ -->
<div class="page" id="page-saldos">
  <div class="chart-card mb">
    <h2>Saldos Individuais — {data['month']}</h2>
    <table><thead><tr><th>Nome</th><th>Saldo Anterior</th><th>Receitas</th><th>Despesas</th><th>Saldo Atual</th></tr></thead>
    <tbody>{bal_rows}</tbody></table>
  </div>
  <div class="chart-card"><h2>Saldo Atual por Cotista</h2><div style="height:440px"><canvas id="saldoChart"></canvas></div></div>
</div>

<script>
const DATA = {json.dumps(data, ensure_ascii=False)};
const CAT_COLORS = {cat_colors_js};
const TYPE_COLORS = {type_colors_js};
const CATS = {json.dumps(CATS)};
const PILL_COLORS = {json.dumps(PILL_COLORS_LIST)};
const COMP_COLORS = ['#4f8ef7','#5dca8a','#f7c948','#f76c6c','#a78bfa','#fb923c','#38bdf8'];
const gridColor = '#252838';
Chart.defaults.color = '#8892a4';
Chart.defaults.font.family = 'Segoe UI, system-ui, Arial, sans-serif';

function brl(v) {{ return 'R$ '+Number(v).toLocaleString('pt-BR',{{minimumFractionDigits:2,maximumFractionDigits:2}}); }}
function brlk(v) {{ return 'R$ '+(v/1000).toFixed(1)+'k'; }}

// ── Tabs ──────────────────────────────────────────────────────────────────────
let mapInit=false, metInit=false;
function switchTab(id) {{
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  const idx=['overview','admin','mapa','vacancia','metragem','estimativa','saldos'].indexOf(id);
  document.querySelectorAll('.tab')[idx].classList.add('active');
  document.getElementById('page-'+id).classList.add('active');
  if(id==='mapa'&&!mapInit){{initMap();mapInit=true;}}
  if(id==='metragem'&&!metInit){{initMetCharts();metInit=true;}}
}}

// ── Multi-select filter ───────────────────────────────────────────────────────
let activeCats = new Set(CATS);

function refreshPills() {{
  const allActive = activeCats.size===CATS.length;
  const allPill = document.getElementById('pill-all');
  allPill.classList.toggle('active', allActive);
  allPill.style.background = allActive ? 'var(--accent)' : '';
  allPill.style.borderColor = allActive ? 'var(--accent)' : '';
  allPill.style.color = allActive ? '#fff' : '';
  CATS.forEach((c,i)=>{{
    const p = document.getElementById('pill-'+i);
    const on = activeCats.has(c);
    p.classList.toggle('active', on);
    p.style.background = on ? PILL_COLORS[i] : '';
    p.style.borderColor = on ? PILL_COLORS[i] : '';
    p.style.color = on ? '#fff' : '';
  }});
}}

function toggleAll() {{
  activeCats = new Set(CATS);
  refreshPills(); updateOverview();
}}

function toggleCat(i) {{
  const cat = CATS[i];
  if(activeCats.has(cat)) {{
    if(activeCats.size===1) return; // keep at least one
    activeCats.delete(cat);
  }} else {{
    activeCats.add(cat);
  }}
  refreshPills(); updateOverview();
}}

function filteredProps() {{ return DATA.properties.filter(p=>activeCats.has(p.cat)); }}

function updateOverview() {{
  const fps = filteredProps();
  const aluguel    = fps.reduce((s,p)=>s+p.aluguel,0);
  const receita    = fps.reduce((s,p)=>s+p.total,0);
  const honorarios = fps.reduce((s,p)=>s+p.honorarios,0);
  const occProps   = fps.filter(p=>p.occupied);
  const occArea    = occProps.reduce((s,p)=>s+p.area,0);
  const m2         = occArea>0 ? aluguel/occArea : 0;
  document.getElementById('dk-aluguel').textContent    = brl(aluguel);
  document.getElementById('dk-receita').textContent    = brl(receita);
  document.getElementById('dk-honorarios').textContent = brl(honorarios);
  document.getElementById('dk-m2').textContent         = 'R$ '+m2.toFixed(2)+'/m²';
  document.getElementById('dk-ocupadas').textContent   = occProps.length+' / '+fps.length;
  document.getElementById('dk-vagas').textContent      = (fps.length-occProps.length)+' un.';

  const selCats=[...activeCats];
  catAluguelChart.data.labels=selCats;
  catAluguelChart.data.datasets[0].data=selCats.map(c=>fps.filter(p=>p.cat===c).reduce((s,p)=>s+p.aluguel,0));
  catAluguelChart.data.datasets[0].backgroundColor=selCats.map(c=>CAT_COLORS[c]+'BB');
  catAluguelChart.data.datasets[0].borderColor=selCats.map(c=>CAT_COLORS[c]);
  catAluguelChart.update();

  const compKeys=['Aluguel','Encargos','Acordo','Condominio','Iptu','Honorarios','Outros'];
  const compMap={{'Aluguel':'aluguel','Encargos':'encargos','Acordo':'acordo','Condominio':'condominio','Iptu':'iptu','Honorarios':'honorarios','Outros':'outros'}};
  const compVals=compKeys.map(k=>fps.reduce((s,p)=>s+(p[compMap[k]]||0),0));
  const total=compVals.reduce((s,v)=>s+v,0);
  const fk=compKeys.filter((_,i)=>compVals[i]>0);
  const fv=compVals.filter(v=>v>0);
  compChart.data.labels=fk.map((k,i)=>k+' ('+((fv[i]/total)*100).toFixed(1)+'%)');
  compChart.data.datasets[0].data=fv;
  compChart.update();

  const sorted=[...fps].filter(p=>p.total>0).sort((a,b)=>b.total-a.total).slice(0,20);
  const maxT=sorted[0]?.total||1;
  const totR=fps.reduce((s,p)=>s+p.total,0)||1;
  document.getElementById('top-table-body').innerHTML=sorted.map(p=>{{
    const col=CAT_COLORS[p.cat]||'#888';
    return `<tr>
      <td>${{p.name}}</td>
      <td><span class="tag" style="background:${{col}}22;color:${{col}}">${{p.cat}}</span></td>
      <td>R$ ${{p.aluguel.toLocaleString('pt-BR',{{minimumFractionDigits:2}})}}</td>
      <td>R$ ${{p.total.toLocaleString('pt-BR',{{minimumFractionDigits:2}})}}</td>
      <td><div class="pct-label">${{(p.total/totR*100).toFixed(1)}}%</div>
          <div class="bar"><div class="bar-fill" style="width:${{(p.total/maxT*100).toFixed(1)}}%;background:${{col}}"></div></div></td>
    </tr>`;
  }}).join('');
}}

// ── Charts: overview ──────────────────────────────────────────────────────────
const catAluguelChart = new Chart(document.getElementById('catAluguelChart'),{{
  type:'bar',
  data:{{ labels:CATS, datasets:[{{label:'Aluguel (R$)',data:CATS.map(c=>DATA.cat_data[c].aluguel),
    backgroundColor:CATS.map(c=>CAT_COLORS[c]+'BB'),borderColor:CATS.map(c=>CAT_COLORS[c]),borderWidth:1,borderRadius:6}}] }},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+brl(ctx.parsed.y)}}}}}},
    scales:{{x:{{ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},
             y:{{ticks:{{color:'#8892a4',callback:v=>brlk(v)}},grid:{{color:gridColor}}}}}}}}
}});
const allCompKeys=Object.keys(DATA.comp_totals), allCompVals=Object.values(DATA.comp_totals);
const totalComp=allCompVals.reduce((s,v)=>s+v,0);
const compChart=new Chart(document.getElementById('compChart'),{{
  type:'doughnut',
  data:{{labels:allCompKeys.map((k,i)=>k+' ('+((allCompVals[i]/totalComp)*100).toFixed(1)+'%)'),
    datasets:[{{data:allCompVals,backgroundColor:COMP_COLORS,borderColor:'#141720',borderWidth:3}}]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{position:'right',labels:{{color:'#e2e8f0',font:{{size:10}},padding:10,boxWidth:10}}}},
      tooltip:{{callbacks:{{label:ctx=>{{const t=ctx.dataset.data.reduce((s,v)=>s+v,0);return' '+brl(ctx.parsed)+' ('+((ctx.parsed/t)*100).toFixed(1)+'%)';}}}}}}}}}}
}});

// ── Charts: admin ─────────────────────────────────────────────────────────────
new Chart(document.getElementById('adminPctChart'),{{type:'bar',
  data:{{labels:CATS,datasets:[{{label:'%',data:CATS.map(c=>DATA.cat_data[c].pct_admin),
    backgroundColor:CATS.map(c=>CAT_COLORS[c]+'BB'),borderColor:CATS.map(c=>CAT_COLORS[c]),borderWidth:1,borderRadius:6}}]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+ctx.parsed.y.toFixed(2)+'%'}}}}}},
    scales:{{x:{{ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},
             y:{{min:0,max:12,ticks:{{color:'#8892a4',callback:v=>v+'%'}},grid:{{color:gridColor}}}}}}}}
}});
new Chart(document.getElementById('adminAbsChart'),{{type:'bar',
  data:{{labels:CATS,datasets:[
    {{label:'Aluguel',data:CATS.map(c=>DATA.cat_data[c].aluguel),backgroundColor:'#4f8ef766',borderColor:'#4f8ef7',borderWidth:1,borderRadius:4}},
    {{label:'Honorários',data:CATS.map(c=>DATA.cat_data[c].honorarios),backgroundColor:'#fb923c77',borderColor:'#fb923c',borderWidth:1,borderRadius:4}}]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{labels:{{color:'#e2e8f0'}}}},tooltip:{{callbacks:{{label:ctx=>' '+brl(ctx.parsed.y)}}}}}},
    scales:{{x:{{ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},
             y:{{ticks:{{color:'#8892a4',callback:v=>brlk(v)}},grid:{{color:gridColor}}}}}}}}
}});

// ── Charts: vacância ──────────────────────────────────────────────────────────
const vtypes=DATA.vacancy_by_type.sort((a,b)=>b.total-a.total);
new Chart(document.getElementById('vaccTypeChart'),{{type:'bar',
  data:{{labels:vtypes.map(v=>v.type),datasets:[
    {{label:'Ocupadas',data:vtypes.map(v=>v.occupied),backgroundColor:'#5dca8a99',borderColor:'#5dca8a',borderWidth:1,borderRadius:4}},
    {{label:'Vagas',data:vtypes.map(v=>v.vacant),backgroundColor:'#f76c6c99',borderColor:'#f76c6c',borderWidth:1,borderRadius:4}}]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{labels:{{color:'#e2e8f0'}}}},tooltip:{{mode:'index'}}}},
    scales:{{x:{{stacked:true,ticks:{{color:'#8892a4'}},grid:{{color:gridColor}}}},y:{{stacked:true,ticks:{{color:'#8892a4'}},grid:{{color:gridColor}}}}}}}}
}});
new Chart(document.getElementById('vaccDonut'),{{type:'doughnut',
  data:{{labels:['Ocupadas','Vagas'],datasets:[{{data:[DATA.barao_occupied_units,DATA.barao_vacant_units],backgroundColor:['#5dca8a','#f76c6c'],borderColor:'#141720',borderWidth:3}}]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{position:'bottom',labels:{{color:'#e2e8f0',padding:14}}}},
      tooltip:{{callbacks:{{label:ctx=>' '+ctx.parsed+' unid. ('+((ctx.parsed/DATA.barao_total_units)*100).toFixed(1)+'%)'}}}}}}}}
}});

// ── Charts: metragem (lazy-init) ──────────────────────────────────────────────
function initMetCharts() {{
  const catM2  = CATS.map(c=>DATA.cat_data[c].aluguel_m2||0);
  const catArea = CATS.map(c=>DATA.cat_data[c].area||0);
  const catOccA = CATS.map(c=>DATA.cat_data[c].occupied_area||0);
  const catVacA = CATS.map(c=>DATA.cat_data[c].vacant_area||0);

  new Chart(document.getElementById('metCatM2Chart'),{{type:'bar',
    data:{{labels:CATS,datasets:[{{label:'R$/m²',data:catM2,
      backgroundColor:CATS.map(c=>CAT_COLORS[c]+'BB'),borderColor:CATS.map(c=>CAT_COLORS[c]),borderWidth:1,borderRadius:6}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' R$ '+ctx.parsed.y.toFixed(2)+'/m²'}}}}}},
      scales:{{x:{{ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},
               y:{{ticks:{{color:'#8892a4',callback:v=>'R$ '+v+'/m²'}},grid:{{color:gridColor}}}}}}}}
  }});

  new Chart(document.getElementById('metAreaDonut'),{{type:'doughnut',
    data:{{labels:CATS.map((c,i)=>c+' ('+((catArea[i]/DATA.total_area)*100).toFixed(1)+'%)'),
      datasets:[{{data:catArea,backgroundColor:CATS.map(c=>CAT_COLORS[c]),borderColor:'#141720',borderWidth:3}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'right',labels:{{color:'#e2e8f0',font:{{size:10}},padding:10,boxWidth:10}}}},
        tooltip:{{callbacks:{{label:ctx=>ctx.parsed.toLocaleString('pt-BR')+' m² ('+((ctx.parsed/DATA.total_area)*100).toFixed(1)+'%)'}}}}}}}}
  }});

  new Chart(document.getElementById('metOccAreaChart'),{{type:'bar',
    data:{{labels:CATS,datasets:[
      {{label:'Área Ocupada (m²)',data:catOccA,backgroundColor:'#5dca8a99',borderColor:'#5dca8a',borderWidth:1,borderRadius:4}},
      {{label:'Área Vaga (m²)',data:catVacA,backgroundColor:'#f76c6c99',borderColor:'#f76c6c',borderWidth:1,borderRadius:4}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{labels:{{color:'#e2e8f0'}}}},tooltip:{{mode:'index',callbacks:{{label:ctx=>ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('pt-BR')+' m²'}}}}}},
      scales:{{x:{{stacked:true,ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},
               y:{{stacked:true,ticks:{{color:'#8892a4',callback:v=>v.toLocaleString('pt-BR')+' m²'}},grid:{{color:gridColor}}}}}}}}
  }});

  const tmKeys=Object.keys(DATA.type_met).sort((a,b)=>DATA.type_met[b].area-DATA.type_met[a].area);
  const tmM2=tmKeys.map(t=>DATA.type_met[t].aluguel_m2||0);
  new Chart(document.getElementById('metTypeM2Chart'),{{type:'bar',
    data:{{labels:tmKeys,datasets:[{{label:'R$/m²',data:tmM2,
      backgroundColor:tmKeys.map(t=>TYPE_COLORS[t]+'BB'||'#4f8ef7BB'),
      borderColor:tmKeys.map(t=>TYPE_COLORS[t]||'#4f8ef7'),borderWidth:1,borderRadius:6}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' R$ '+ctx.parsed.y.toFixed(2)+'/m²'}}}}}},
      scales:{{x:{{ticks:{{color:'#8892a4'}},grid:{{color:gridColor}}}},
               y:{{ticks:{{color:'#8892a4',callback:v=>'R$ '+v}},grid:{{color:gridColor}}}}}}}}
  }});
}}

// ── Estimativa ────────────────────────────────────────────────────────────────
// ── Estimativa — keyed by "cat||type" combos ──────────────────────────────────
const CAT_SHORT = {{
  "Rua Florêncio de Abreu, 60":       "Fl. Abreu 60",
  "Rua Florêncio de Abreu, 66/70/74": "Fl. Abreu 66+",
  "Rua 25 de Março, 669":             "25 Março 669",
  "Rua 25 de Março, 661":             "25 Março 661",
  "Ed. Barão":                        "Barão",
}};

const estPrices = {{}};  // ct_key -> R$/m²
const estOccPct = {{}};  // ct_key -> occupancy % (0-100)
let estCatChart, estDiffChart;

// Sorted ct_keys: Rua properties first, then Barão by area desc
const CT_KEYS = Object.keys(DATA.cat_type).sort((a,b)=>{{
  const da=DATA.cat_type[a], db=DATA.cat_type[b];
  if(da.cat!==db.cat) return CATS.indexOf(da.cat)-CATS.indexOf(db.cat);
  return db.area-da.area;
}});

function ctLabel(key) {{
  const d=DATA.cat_type[key];
  return (CAT_SHORT[d.cat]||d.cat)+' · '+d.type;
}}

function buildEstInputs() {{
  const priceDiv=document.getElementById('est-price-inputs');
  const occDiv  =document.getElementById('est-occ-inputs');

  priceDiv.innerHTML=CT_KEYS.map(key=>{{
    const d=DATA.cat_type[key];
    const def=d.aluguel_m2||0;
    estPrices[key]=def;
    const color=CAT_COLORS[d.cat]||'#4f8ef7';
    const safeKey=key.replace(/[^a-z0-9]/gi,'_');
    return `<div class="est-row">
      <div class="est-type-dot" style="background:${{color}}"></div>
      <span class="est-label">${{ctLabel(key)}}</span>
      <input class="est-input" type="number" id="ep_${{safeKey}}" value="${{def.toFixed(2)}}" step="0.50"
        onchange="estPrices['${{key}}']=parseFloat(this.value)||0;recalcEstimativa()"/>
      <span class="est-unit">/m²</span>
    </div>`;
  }}).join('');

  occDiv.innerHTML=CT_KEYS.map(key=>{{
    const d=DATA.cat_type[key];
    const def=d.units>0?Math.round(d.occupied/d.units*100):100;
    estOccPct[key]=def;
    const color=CAT_COLORS[d.cat]||'#4f8ef7';
    const safeKey=key.replace(/[^a-z0-9]/gi,'_');
    return `<div class="est-row">
      <div class="est-type-dot" style="background:${{color}}"></div>
      <span class="est-label">${{ctLabel(key)}}</span>
      <input class="est-slider" type="range" min="0" max="100" value="${{def}}"
        oninput="estOccPct['${{key}}']=parseInt(this.value);document.getElementById('pct_${{safeKey}}').textContent=this.value+'%';recalcEstimativa()"/>
      <span class="est-pct" id="pct_${{safeKey}}">${{def}}%</span>
      <span style="font-size:.65rem;color:var(--muted2);margin-left:4px">${{d.occupied}}/${{d.units}} atual</span>
    </div>`;
  }}).join('');
}}

function recalcEstimativa() {{
  let totalEst=0, totalArea=0, totalUnits=0;
  const catEst={{}};
  CATS.forEach(c=>{{catEst[c]={{est:0,units:0,area:0}}}});

  // Per ct_key: sort units by descending aluguel, occupy top N
  CT_KEYS.forEach(key=>{{
    const props=DATA.properties.filter(p=>p.ct_key===key);
    const nOcc=Math.round(props.length*(estOccPct[key]||0)/100);
    const sorted=[...props].sort((a,b)=>b.aluguel-a.aluguel);
    sorted.forEach((p,i)=>{{
      const will=i<nOcc;
      const price=estPrices[key]||0;
      const est=will?price*p.area:0;
      catEst[p.cat].est+=est;
      if(will){{catEst[p.cat].units++;catEst[p.cat].area+=p.area;}}
      totalEst+=est;
      if(will){{totalArea+=p.area;totalUnits++;}}
    }});
  }});

  const diff=totalEst-DATA.total_aluguel;
  document.getElementById('er-aluguel').textContent=brl(totalEst);
  document.getElementById('er-diff').textContent=(diff>=0?'+':'')+brl(diff);
  document.getElementById('er-diff').style.color=diff>=0?'var(--green)':'var(--red)';
  document.getElementById('er-area').textContent=totalArea.toLocaleString('pt-BR')+' m²';
  document.getElementById('er-units').textContent=totalUnits+' unidades';

  const estVals=CATS.map(c=>catEst[c].est);
  const curVals=CATS.map(c=>DATA.cat_data[c].aluguel);
  if(estCatChart){{
    estCatChart.data.datasets[1].data=estVals;
    estCatChart.update();
    estDiffChart.data.datasets[0].data=CATS.map((c,i)=>estVals[i]-curVals[i]);
    estDiffChart.data.datasets[0].backgroundColor=CATS.map((c,i)=>(estVals[i]-curVals[i])>=0?'#5dca8a88':'#f76c6c88');
    estDiffChart.data.datasets[0].borderColor=CATS.map((c,i)=>(estVals[i]-curVals[i])>=0?'#5dca8a':'#f76c6c');
    estDiffChart.update();
  }}
}}

function initEstimativa() {{
  buildEstInputs();
  estCatChart=new Chart(document.getElementById('estCatChart'),{{type:'bar',
    data:{{labels:CATS,datasets:[
      {{label:'Atual',data:CATS.map(c=>DATA.cat_data[c].aluguel),backgroundColor:CATS.map(c=>CAT_COLORS[c]+'44'),borderColor:CATS.map(c=>CAT_COLORS[c]),borderWidth:1,borderRadius:4}},
      {{label:'Estimado',data:CATS.map(c=>DATA.cat_data[c].aluguel),backgroundColor:CATS.map(c=>CAT_COLORS[c]+'99'),borderColor:CATS.map(c=>CAT_COLORS[c]),borderWidth:2,borderRadius:4}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{labels:{{color:'#e2e8f0'}}}},tooltip:{{callbacks:{{label:ctx=>' '+brl(ctx.parsed.y)}}}}}},
      scales:{{x:{{ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},y:{{ticks:{{color:'#8892a4',callback:v=>brlk(v)}},grid:{{color:gridColor}}}}}}}}
  }});
  estDiffChart=new Chart(document.getElementById('estDiffChart'),{{type:'bar',
    data:{{labels:CATS,datasets:[{{label:'Δ vs Atual',data:CATS.map(()=>0),
      backgroundColor:'#5dca8a88',borderColor:'#5dca8a',borderWidth:1,borderRadius:5}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>(ctx.parsed.y>=0?'+':'')+brl(ctx.parsed.y)}}}}}},
      scales:{{x:{{ticks:{{color:'#8892a4',maxRotation:20}},grid:{{color:gridColor}}}},
               y:{{ticks:{{color:'#8892a4',callback:v=>(v>=0?'+':'')+brlk(v)}},grid:{{color:gridColor}}}}}}}}
  }});
  recalcEstimativa();
}}

// ── Charts: saldos ────────────────────────────────────────────────────────────
const bals=DATA.balances.filter(b=>b.atual!==0).sort((a,b)=>b.atual-a.atual);
new Chart(document.getElementById('saldoChart'),{{type:'bar',
  data:{{labels:bals.map(b=>b.name.split(' ').slice(0,2).join(' ')),
    datasets:[{{label:'Saldo Atual',data:bals.map(b=>b.atual),
      backgroundColor:bals.map(b=>b.atual>=0?'#5dca8a77':'#f76c6c77'),
      borderColor:bals.map(b=>b.atual>=0?'#5dca8a':'#f76c6c'),borderWidth:1,borderRadius:5}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+brl(ctx.parsed.x)}}}}}},
    scales:{{x:{{ticks:{{color:'#8892a4',callback:v=>brlk(v)}},grid:{{color:gridColor}}}},
             y:{{ticks:{{color:'#e2e8f0',font:{{size:11}}}},grid:{{color:gridColor}}}}}}}}
}});

// ── Map ───────────────────────────────────────────────────────────────────────
function initMap() {{
  const markers = {map_markers_js};
  const map=L.map('map').setView([-23.5450,-46.6335],15);
  L.tileLayer('https://{{s}}.basemaps.cartocdn.com/dark_all/{{z}}/{{x}}/{{y}}{{r}}.png',{{attribution:'&copy; OpenStreetMap &amp; CartoDB',maxZoom:19}}).addTo(map);
  const maxAl=Math.max(...markers.map(m=>m.aluguel));
  markers.forEach(m=>{{
    const r=Math.max(18,Math.min(60,(m.aluguel/maxAl)*60));
    const color=CAT_COLORS[m.cat]||'#4f8ef7';
    const circle=L.circleMarker([m.lat,m.lng],{{radius:r,fillColor:color,color:color,weight:2,opacity:.9,fillOpacity:.3}}).addTo(map);
    circle.bindPopup(`<strong>${{m.cat}}</strong>
      <div style="color:#8892a4;font-size:.72rem;margin:4px 0 8px">${{m.addr}}</div>
      <div class="pop-row"><span>Aluguel:</span><span>${{brl(m.aluguel)}}</span></div>
      <div class="pop-row"><span>Receita Total:</span><span>${{brl(m.total)}}</span></div>
      <div class="pop-row"><span>Unidades:</span><span>${{m.occupied}}/${{m.units}} ocupadas</span></div>`);
  }});
}}

// ── Init ──────────────────────────────────────────────────────────────────────
refreshPills();
updateOverview();

// Init estimativa when tab first opened
const estTab = document.querySelectorAll('.tab')[5];
let estInit = false;
estTab.addEventListener('click', ()=>{{ if(!estInit){{ initEstimativa(); estInit=true; }} }});
</script>
</body>
</html>"""

OUTPUT_PATH.write_text(html, encoding="utf-8")
print(f"Dashboard gerado: {OUTPUT_PATH}")
print(f"Localidades mapeadas: {len(data['locations'])}")
