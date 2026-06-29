"""
generar.py — BRE-B Tablero de Tesorería · Febor Entidad Cooperativa
====================================================================
Uso:
    python generar.py

Lee  : datos/datos.xlsx   (hoja 1767243600_informe_movimientos)
Genera: docs/index.html   (tablero listo para GitHub Pages)

Cada vez que actualices el Excel, vuelve a correr este script
y luego haz git push para publicar los cambios.
"""

import pandas as pd
import numpy as np
import json
import math
import os
from pathlib import Path
from collections import defaultdict
from datetime import datetime

# ── Rutas ─────────────────────────────────────────────────────────────────────
BASE = Path(__file__).parent
EXCEL_PATH    = BASE / "datos" / "datos.xlsx"
TEMPLATE_PATH = BASE / "tablero_template.html"
OUTPUT_PATH   = BASE / "docs" / "index.html"

# ── Parámetros del modelo ─────────────────────────────────────────────────────
HOJA        = "1767243600_informe_movimientos"
Z_NC95      = 1.64          # factor NC 95%
LIM_REG     = 10_000_000   # límite regulatorio (10M COP)
HIST_DIAS   = 60            # días de historial para gráfica proyección

# ── 1. Carga de datos ─────────────────────────────────────────────────────────
print("📂  Leyendo Excel…")
df = pd.read_excel(EXCEL_PATH, sheet_name=HOJA, engine="openpyxl")

df["trx_dt"]  = pd.to_datetime(df["trx_fecha"], errors="coerce")
df["Valor2"]  = pd.to_numeric(df["Valor"],     errors="coerce").fillna(0)
df["Mes"]     = df["trx_dt"].dt.strftime("%Y-%m")
df["fecha"]   = df["trx_dt"].dt.strftime("%Y-%m-%d")
df["hora"]    = df["trx_dt"].dt.hour

def cat_dia(row):
    if str(row.get("Festivo o No","")) == "Festivos":
        return "Festivo"
    if str(row.get("Tipo Día","")) == "Fin de Semana":
        return "Fin de semana"
    return "Día hábil"

def franja(h):
    if  0 <= h < 6:  return "Madrugada"
    if  6 <= h < 12: return "Mañana"
    if 12 <= h < 18: return "Tarde"
    return "Noche"

df["cat_dia"] = df.apply(cat_dia, axis=1)
df["franja"]  = df["hora"].apply(franja)
df["estado"]  = df["Nombre Error"].fillna("Desconocido").astype(str)
df["doc"]     = df["Documento"].fillna("").astype(str)
df["nombre"]  = df["Nombre Cliente Origen"].fillna("").astype(str)

print(f"    {len(df):,} filas | {df['Mes'].nunique()} meses | {df['doc'].nunique():,} asociados")

# ── 2. Índices para DATA_RAW (compact array) ──────────────────────────────────
fechas_uniq  = sorted(df["fecha"].unique())
estados_uniq = sorted(df["estado"].unique())
meses_uniq   = sorted(df["Mes"].unique())
cats_uniq    = ["Día hábil", "Fin de semana", "Festivo"]
franjas_uniq = ["Madrugada", "Mañana", "Tarde", "Noche"]
docs_uniq    = sorted(df["doc"].unique())
nombres_uniq = list(df.groupby("doc")["nombre"].first().reindex(docs_uniq).fillna(""))

fi = {v: i for i,v in enumerate(fechas_uniq)}
ei = {v: i for i,v in enumerate(estados_uniq)}
mi = {v: i for i,v in enumerate(meses_uniq)}
ci = {v: i for i,v in enumerate(cats_uniq)}
fri= {v: i for i,v in enumerate(franjas_uniq)}
di = {v: i for i,v in enumerate(docs_uniq)}

rows = []
for _, r in df.iterrows():
    fe = fi.get(r["fecha"],   -1)
    e  = ei.get(r["estado"],  -1)
    m  = mi.get(r["Mes"],     -1)
    c  = ci.get(r["cat_dia"], -1)
    fr = fri.get(r["franja"], -1)
    d  = di.get(r["doc"],     -1)
    n  = di.get(r["doc"],     -1)   # mismo índice que doc para nombres_uniq
    if fe < 0 or e < 0: continue
    rows.append([fe, e, m, c, fr, r["hora"], round(r["Valor2"]/1e6, 4),
                 d, n])

DATA_RAW = {
    "fechas":  fechas_uniq,
    "estados": estados_uniq,
    "meses":   meses_uniq,
    "cats":    cats_uniq,
    "franjas": franjas_uniq,
    "docs":    docs_uniq,
    "nombres": nombres_uniq,
    "rows":    rows,
}

# ── 3. CUBE (agregado por estado|mes|cat|franja) ──────────────────────────────
print("🔲  Construyendo CUBE…")

# Outlier detection (solo Éxito)
exito = df[df["estado"] == "Exito"]["Valor2"]
q1, q3 = exito.quantile(0.25), exito.quantile(0.75)
iqr    = q3 - q1
lim_iqr = q3 + 1.5 * iqr

def is_out(v, estado):
    if estado != "Exito": return False
    return v > lim_iqr or v >= LIM_REG

ESTADOS_LIST = estados_uniq + ["Todos"]
MESES_LIST   = meses_uniq   + ["Todos"]
CATS_LIST    = cats_uniq    + ["Todos"]
FRANJAS_LIST = franjas_uniq + ["Todos"]

def make_cell():
    return {"n":0,"m":0,"docs":set(),"outs":0,
            "hm":{},"mm":{},"by_cat":{},"by_franja":{},"top":{}}

CUBE = {}

for _, r in df.iterrows():
    est = r["estado"];  mes = r["Mes"]
    cat = r["cat_dia"]; fr  = r["franja"]
    val = r["Valor2"];  hora= r["hora"]
    doc = r["doc"];     nom = r["nombre"]
    out = is_out(val, est)

    for eb in [est, "Todos"]:
        for mb in [mes, "Todos"]:
            for cb in [cat, "Todos"]:
                for fb in [fr, "Todos"]:
                    k = f"{eb}|{mb}|{cb}|{fb}"
                    if k not in CUBE: CUBE[k] = make_cell()
                    c = CUBE[k]
                    c["n"] += 1
                    c["m"] += val
                    c["docs"].add(doc)
                    if out: c["outs"] += 1
                    c["hm"].setdefault(hora, {"n":0,"m":0})
                    c["hm"][hora]["n"] += 1
                    c["hm"][hora]["m"] += val
                    c["mm"].setdefault(mes, {"n":0,"m":0})
                    c["mm"][mes]["n"] += 1
                    c["mm"][mes]["m"] += val
                    c["by_cat"].setdefault(cat, {"n":0,"m":0})
                    c["by_cat"][cat]["n"] += 1
                    c["by_cat"][cat]["m"] += val
                    c["by_franja"].setdefault(fr, {"n":0,"m":0})
                    c["by_franja"][fr]["n"] += 1
                    c["by_franja"][fr]["m"] += val
                    if doc not in c["top"]:
                        c["top"][doc] = {"d":doc,"n":nom,"ops":0,"m":0}
                    c["top"][doc]["ops"] += 1
                    c["top"][doc]["m"]   += val

# Serializar CUBE
MES_ORDER = sorted(meses_uniq)
CUBE_JSON = {}
for k, c in CUBE.items():
    by_hora   = [{"h":h,"n":v["n"],"m":round(v["m"])}
                 for h,v in sorted(c["hm"].items())]
    by_mes    = [{"m":m,"n":v["n"],"v":round(v["m"])}
                 for m,v in sorted(c["mm"].items())]
    by_cat    = {cat: {"n":v["n"],"m":round(v["m"])}
                 for cat,v in c["by_cat"].items()}
    by_franja = {fr:  {"n":v["n"],"m":round(v["m"])}
                 for fr,v in c["by_franja"].items()}
    top_docs  = sorted(c["top"].values(),
                       key=lambda x: -x["ops"])[:10]
    top_docs  = [{"d":t["d"],"n":t["n"],
                  "ops":t["ops"],"m":round(t["m"])} for t in top_docs]
    CUBE_JSON[k] = {
        "kpis": {"n_ops":c["n"],"monto_total":round(c["m"]),
                 "monto_prom": round(c["m"]/c["n"]) if c["n"] else 0,
                 "docs":len(c["docs"]),"outliers":c["outs"]},
        "by_hora": by_hora,
        "by_mes":  by_mes,
        "by_cat":  by_cat,
        "by_franja": by_franja,
        "top_docs":  top_docs,
    }

print(f"    {len(CUBE_JSON):,} celdas en CUBE")

# ── 4. DIA_CUBE (desglose por día de semana) ──────────────────────────────────
print("📅  Construyendo DIA_CUBE…")
DIAS_HAB  = ["Lunes","Martes","Miércoles","Jueves","Viernes"]
DIAS_FIND = ["Sábado","Domingo"]
DIA_NAMES = ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"]

df["dia_semana"] = df["trx_dt"].dt.dayofweek.map(
    {0:"Lunes",1:"Martes",2:"Miércoles",3:"Jueves",
     4:"Viernes",5:"Sábado",6:"Domingo"})

DIA_CUBE = {}

for _, r in df.iterrows():
    est = r["estado"];  mes = r["Mes"]
    cat = r["cat_dia"]; fr  = r["franja"]
    val = r["Valor2"];  dia = r["dia_semana"]
    doc = r["doc"];     fe  = r["fecha"]

    for eb in [est, "Todos"]:
        for mb in [mes, "Todos"]:
            for cb in [cat, "Todos"]:
                for fb in [fr, "Todos"]:
                    k = f"{eb}|{mb}|{cb}|{fb}"
                    if k not in DIA_CUBE:
                        DIA_CUBE[k] = {"hab":{},"finde":{},"festivo":{}}
                    dc = DIA_CUBE[k]

                    if cat == "Día hábil":
                        dc["hab"].setdefault(dia, {"n":0,"m":0,"docs":set()})
                        dc["hab"][dia]["n"] += 1
                        dc["hab"][dia]["m"] += val
                        dc["hab"][dia]["docs"].add(doc)
                    elif cat == "Fin de semana":
                        dc["finde"].setdefault(dia, {"n":0,"m":0,"docs":set()})
                        dc["finde"][dia]["n"] += 1
                        dc["finde"][dia]["m"] += val
                        dc["finde"][dia]["docs"].add(doc)
                    else:
                        MN = ["Ene","Feb","Mar","Abr","May","Jun",
                              "Jul","Ago","Sep","Oct","Nov","Dic"]
                        d_obj = datetime.strptime(fe, "%Y-%m-%d")
                        lbl = f"{d_obj.day} {MN[d_obj.month-1]}"
                        dc["festivo"].setdefault(fe, {"d":lbl,"n":0,"m":0,"docs":set()})
                        dc["festivo"][fe]["n"] += 1
                        dc["festivo"][fe]["m"] += val
                        dc["festivo"][fe]["docs"].add(doc)

DIA_CUBE_JSON = {}
for k, dc in DIA_CUBE.items():
    hab  = [{"d":d,"n":v["n"],"m":round(v["m"]),"docs":len(v["docs"])}
            for d,v in dc["hab"].items()
            if d in DIAS_HAB]
    hab.sort(key=lambda x: DIAS_HAB.index(x["d"]))
    finde= [{"d":d,"n":v["n"],"m":round(v["m"]),"docs":len(v["docs"])}
            for d,v in dc["finde"].items()
            if d in DIAS_FIND]
    finde.sort(key=lambda x: DIAS_FIND.index(x["d"]))
    fest = sorted(
        [{"d":v["d"],"n":v["n"],"m":round(v["m"]),"docs":len(v["docs"])}
         for v in dc["festivo"].values()],
        key=lambda x: x["d"])
    DIA_CUBE_JSON[k] = {"hab":hab,"finde":finde,"festivo":fest}

print(f"    {len(DIA_CUBE_JSON):,} celdas en DIA_CUBE")

# ── 5. DOCMAP y ASOC_OUT ──────────────────────────────────────────────────────
print("👥  Construyendo DOCMAP / ASOC_OUT…")

exito_df = df[df["estado"] == "Exito"]

DOCMAP = {}
for doc, grp in exito_df.groupby("doc"):
    dias_hab  = grp[grp["cat_dia"]=="Día hábil"]["fecha"].nunique()
    dias_find = grp[grp["cat_dia"]=="Fin de semana"]["fecha"].nunique()
    dias_fest = grp[grp["cat_dia"]=="Festivo"]["fecha"].nunique()
    DOCMAP[doc] = {
        "nom": grp["nombre"].iloc[0],
        "ops": len(grp),
        "m":   round(grp["Valor2"].sum()),
        "bm":  round(grp[grp["franja"]=="Mañana"]["Valor2"].sum()),
        "bh":  round(grp[grp["franja"]=="Tarde"]["Valor2"].sum()),
        "bc":  dias_hab,
        "bf":  dias_find,
    }

# NOMBRES ordenados
NOMBRES = sorted([
    {"d": doc, "n": info["nom"]}
    for doc, info in DOCMAP.items()
    if info["nom"].strip()
], key=lambda x: x["n"])

# Outliers por asociado
lim_iqr_perc = exito_df.groupby("doc")["Valor2"].quantile(0.95)

ASOC_OUT = {}
for doc, grp in exito_df.groupby("doc"):
    lim_doc = lim_iqr_perc.get(doc, lim_iqr)
    outs = grp[grp["Valor2"] > min(lim_doc, LIM_REG)]
    if len(outs) == 0: continue
    ASOC_OUT[doc] = {
        "nom":  grp["nombre"].iloc[0],
        "n":    len(outs),
        "m":    round(outs["Valor2"].sum()),
        "max":  round(outs["Valor2"].max()),
        "txs":  outs[["fecha","Valor2","franja","cat_dia"]]\
                    .rename(columns={"Valor2":"v","cat_dia":"c","franja":"f"})\
                    .assign(v=lambda x: x["v"].round())\
                    .to_dict("records")[:20]
    }

print(f"    DOCMAP: {len(DOCMAP):,} | ASOC_OUT: {len(ASOC_OUT):,}")

# ── 6. ASOC_DIA ───────────────────────────────────────────────────────────────
ASOC_DIA = {}
for doc, grp in exito_df.groupby("doc"):
    hab  = [{"d":d,"n":int(sg["Valor2"].count()),"m":round(float(sg["Valor2"].sum())),"docs":1}
            for d,sg in grp[grp["cat_dia"]=="Día hábil"].groupby("dia_semana")
            if d in DIAS_HAB]
    hab.sort(key=lambda x: DIAS_HAB.index(x["d"]) if x["d"] in DIAS_HAB else 99)
    finde= [{"d":d,"n":int(sg["Valor2"].count()),"m":round(float(sg["Valor2"].sum())),"docs":1}
            for d,sg in grp[grp["cat_dia"]=="Fin de semana"].groupby("dia_semana")
            if d in DIAS_FIND]
    finde.sort(key=lambda x: DIAS_FIND.index(x["d"]) if x["d"] in DIAS_FIND else 99)
    if hab or finde:
        ASOC_DIA[doc] = {"hab":hab,"finde":finde,"festivo":[]}

# ── 7. PROJ_STATS (estadísticas de proyección) ────────────────────────────────
print("📊  Calculando PROJ_STATS…")

exito_d = exito_df.groupby("fecha")["Valor2"].sum()
mu_dia   = float(exito_d.mean())
sig_dia  = float(exito_d.std())
mu_ops_d = float(exito_df.groupby("fecha").size().mean())
sig_ops  = float(exito_df.groupby("fecha").size().std())
cv_dia   = round(sig_dia / mu_dia * 100, 2) if mu_dia else 0

fi_str = df["fecha"].min()
ff_str = df["fecha"].max()
n_dias = exito_d.shape[0]

hist = exito_d.tail(HIST_DIAS).reset_index()
hist.columns = ["f","m"]
hist["m"] = (hist["m"] / 1e6).round(3)
hist_list = hist.to_dict("records")

PROJ_STATS = {
    "mu_dia":     round(mu_dia),
    "sigma_dia":  round(sig_dia),
    "mu_ops":     round(mu_ops_d, 1),
    "sigma_ops":  round(sig_ops,  1),
    "cv_dia":     cv_dia,
    "min_dia":    round(float(exito_d.min())),
    "max_dia":    round(float(exito_d.max())),
    "fecha_inicio": fi_str,
    "fecha_fin":    ff_str,
    "n_dias":       n_dias,
    "Z":            Z_NC95,
    "fuente":       "BRE-B Enviar · Solo operaciones Éxito",
    "hist":         hist_list,
}

print(f"    μ/día = ${mu_dia/1e6:.1f}M | σ = ${sig_dia/1e6:.1f}M | CV = {cv_dia:.1f}%")
print(f"    Período: {fi_str} → {ff_str} | {n_dias} días")

# ── 8. Generar el bloque de datos JS ─────────────────────────────────────────
print("🔧  Generando bloque JS…")

def js(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",",":"))

# Construir meses order y mesNombres para el select
MES_NOMBRES = {"2025-10":"Oct 2025","2025-11":"Nov 2025","2025-12":"Dic 2025",
               "2026-01":"Ene 2026","2026-02":"Feb 2026","2026-03":"Mar 2026",
               "2026-04":"Abr 2026","2026-05":"May 2026","2026-06":"Jun 2026",
               "2026-07":"Jul 2026","2026-08":"Ago 2026","2026-09":"Sep 2026"}
# Auto-generar para cualquier mes que aparezca
MESES_ESP = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
for m in meses_uniq:
    if m not in MES_NOMBRES:
        y, mo = m.split("-")
        MES_NOMBRES[m] = f"{MESES_ESP[int(mo)-1]} {y}"

# Actualizar hdr-periodo con el rango real
periodo_label = f"{MESES_ESP[int(fi_str[5:7])-1]} {fi_str[:4]} – {MESES_ESP[int(ff_str[5:7])-1]} {ff_str[:4]}"

data_block = f"""<script>
// ── Variables globales ─────────────────────────────────────────────────────────────────
// Generado automáticamente por generar.py el {datetime.now().strftime('%Y-%m-%d %H:%M')}
// Período: {fi_str} al {ff_str} | {len(df):,} operaciones | {df['doc'].nunique():,} asociados
// ──────────────────────────────────────────────────────────────────────────────────────

let CUBE={js(CUBE_JSON)};
let DOCMAP={js(DOCMAP)};
let NOMBRES={js(NOMBRES)};
let ASOC_OUT={js(ASOC_OUT)};
let DIA_CUBE={js(DIA_CUBE_JSON)};
let ASOC_DIA={js(ASOC_DIA)};
let ESTADOS={js(estados_uniq)};
let PROJ_STATS={js(PROJ_STATS)};
let DATA_RAW={js(DATA_RAW)};

// ── Inicialización dinámica ────────────────────────────────────────────────────────────
(function(){{
  const sel=document.getElementById('f-mes');
  if(sel){{
    const mesNombres={js(MES_NOMBRES)};
    {js(sorted(meses_uniq))}.forEach(m=>{{
      if(!m)return;
      const [y,mo]=m.split('-');
      sel.innerHTML+= `<option value="${{m}}">${{mesNombres[m]||m}}</option>`;
    }});
  }}
  // Actualizar header período
  const hp=document.getElementById('hdr-periodo');
  if(hp) hp.textContent='{periodo_label}';
}})();
render();
</script>
"""

# ── 9. Insertar en el template y guardar ─────────────────────────────────────
print("💾  Escribiendo tablero HTML…")

with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
    template = f.read()

if "<!-- DATA_PLACEHOLDER -->" not in template:
    raise ValueError("❌  tablero_template.html no tiene <!-- DATA_PLACEHOLDER -->")

output_html = template.replace("<!-- DATA_PLACEHOLDER -->", data_block)

OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    f.write(output_html)

size_mb = OUTPUT_PATH.stat().st_size / 1e6
print(f"\n✅  Tablero generado: docs/index.html ({size_mb:.1f} MB)")
print(f"   Próximo paso: git add . && git commit -m 'Actualizar datos' && git push")
