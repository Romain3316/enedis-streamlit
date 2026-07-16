import base64
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="CMA | Pré-diagnostic photovoltaïque", page_icon="☀️", layout="wide")

CMA_BLUE = "#172C4C"
CMA_RED = "#E53935"
CMA_LIGHT = "#F4F6F8"
WEEKDAYS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
WEEKDAY_MAP = dict(enumerate(WEEKDAYS))
MONTHS = {1:"Janvier",2:"Février",3:"Mars",4:"Avril",5:"Mai",6:"Juin",7:"Juillet",8:"Août",9:"Septembre",10:"Octobre",11:"Novembre",12:"Décembre"}

st.markdown(f"""
<style>
.block-container {{max-width:1500px;padding-top:1.2rem;padding-bottom:3rem}}
[data-testid="stSidebar"] {{background:{CMA_LIGHT};border-right:1px solid #dde3ea}}
.hero {{display:flex;align-items:center;justify-content:space-between;gap:24px;padding:28px 32px;margin-bottom:20px;background:linear-gradient(110deg,{CMA_BLUE} 0%,{CMA_BLUE} 72%,{CMA_RED} 72%,{CMA_RED} 100%);border-radius:18px;color:white;box-shadow:0 10px 30px rgba(23,44,76,.14)}}
.hero h1 {{margin:0;font-size:2rem;font-weight:800}}
.hero p {{margin:8px 0 0;opacity:.92}}
.logo {{min-width:190px;text-align:center;background:white;color:{CMA_BLUE};border-radius:12px;padding:12px 18px;font-weight:800;font-size:1.15rem}}
.card {{background:white;border:1px solid #e5e9ef;border-radius:16px;padding:20px;box-shadow:0 5px 18px rgba(23,44,76,.06);margin-bottom:18px}}
div[data-testid="stMetric"] {{background:white;border:1px solid #e5e9ef;padding:15px 17px;border-radius:14px;box-shadow:0 4px 14px rgba(23,44,76,.05)}}
.stDownloadButton>button,.stButton>button {{background:{CMA_BLUE};color:white;border:0;border-radius:10px;font-weight:700}}
h2,h3 {{color:{CMA_BLUE}}}
</style>
""", unsafe_allow_html=True)


def img64(path: Path):
    if not path.exists():
        return None
    data = base64.b64encode(path.read_bytes()).decode()
    ext = path.suffix.lower().replace(".", "")
    return f"data:image/{'jpeg' if ext in ('jpg','jpeg') else ext};base64,{data}"


def header():
    candidates = [Path("logo_cma.png"), Path("logo_cma.jpg"), Path("assets/logo_cma.png")]
    uri = next((img64(p) for p in candidates if p.exists()), None)
    logo = f'<div class="logo"><img src="{uri}" style="max-height:74px;max-width:210px"></div>' if uri else '<div class="logo">CMA<br><span style="font-size:.72rem">NOUVELLE-AQUITAINE</span></div>'
    st.markdown(f'<div class="hero"><div><h1>Pré-diagnostic photovoltaïque</h1><p>Analyse automatisée des courbes de charge Enedis</p></div>{logo}</div>', unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def read_file(content: bytes, name: str):
    b = BytesIO(content)
    if name.lower().endswith(".csv"):
        df = None
        for sep in [";", ","]:
            for enc in ["utf-8", "latin-1"]:
                try:
                    b.seek(0)
                    candidate = pd.read_csv(b, sep=sep, encoding=enc)
                    if {"Horodate", "Valeur"}.issubset(candidate.columns):
                        df = candidate
                        break
                except Exception:
                    pass
            if df is not None:
                break
        if df is None:
            raise ValueError("CSV illisible ou colonnes Horodate/Valeur absentes.")
    else:
        df = pd.read_excel(b)
    if not {"Horodate", "Valeur"}.issubset(df.columns):
        raise ValueError("Le fichier doit contenir Horodate et Valeur.")
    cols = [c for c in ["Unité", "Horodate", "Valeur", "Nature", "Pas"] if c in df.columns]
    df = df[cols].copy()
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)
    df["Valeur"] = pd.to_numeric(df["Valeur"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    return df.dropna(subset=["Horodate", "Valeur"]).sort_values("Horodate").reset_index(drop=True)


def detect_step(df):
    d = df["Horodate"].diff()
    d = d[d > pd.Timedelta(0)]
    if d.empty:
        raise ValueError("Pas de temps indétectable.")
    m = d.mode()
    return m.iloc[0] if not m.empty else d.median()


def detect_unit(df):
    if "Unité" not in df.columns:
        return "Wh"
    s = df["Unité"].dropna().astype(str).str.strip()
    s = s[s != ""]
    return s.mode().iloc[0] if not s.empty else "Wh"


def enrich(df, step, unit):
    out = df.copy()
    h = step.total_seconds()/3600
    u = str(unit).lower().replace(" ", "")
    if u in {"wh", "w.h"}:
        out["Energie_kWh"] = out["Valeur"]/1000
        out["Puissance_kW"] = out["Energie_kWh"]/h
        note = "Valeurs interprétées comme une énergie en Wh par intervalle."
    elif u in {"kwh", "kw.h"}:
        out["Energie_kWh"] = out["Valeur"]
        out["Puissance_kW"] = out["Energie_kWh"]/h
        note = "Valeurs interprétées comme une énergie en kWh par intervalle."
    elif u in {"w", "watt", "watts"}:
        out["Puissance_kW"] = out["Valeur"]/1000
        out["Energie_kWh"] = out["Puissance_kW"]*h
        note = "Valeurs interprétées comme une puissance moyenne en W."
    elif u in {"kw", "kilowatt", "kilowatts"}:
        out["Puissance_kW"] = out["Valeur"]
        out["Energie_kWh"] = out["Puissance_kW"]*h
        note = "Valeurs interprétées comme une puissance moyenne en kW."
    else:
        out["Energie_kWh"] = out["Valeur"]/1000
        out["Puissance_kW"] = out["Energie_kWh"]/h
        note = f"Unité « {unit} » non reconnue : hypothèse Wh par intervalle."
    return out, note


def hourly(df):
    x = df.copy(); x["Heure"] = x["Horodate"].dt.floor("h")
    return x.groupby("Heure", as_index=False).agg(Energie_kWh=("Energie_kWh","sum"), Puissance_kW=("Puissance_kW","mean"), Nb_points=("Valeur","size")).rename(columns={"Heure":"Horodate"})


def daily(df):
    x = df.copy(); x["Date"] = x["Horodate"].dt.normalize()
    r = x.groupby("Date", as_index=False).agg(Consommation_kWh=("Energie_kWh","sum"), Puissance_moyenne_kW=("Puissance_kW","mean"), Puissance_max_kW=("Puissance_kW","max"), Nb_points=("Valeur","size"))
    r["Jour_num"] = r["Date"].dt.weekday; r["Jour"] = r["Jour_num"].map(WEEKDAY_MAP); r["Année"] = r["Date"].dt.year; r["Mois"] = r["Date"].dt.month.map(MONTHS)
    return r


def monthly(df):
    x = df.copy(); x["Mois_date"] = x["Horodate"].dt.to_period("M").dt.to_timestamp()
    return x.groupby("Mois_date", as_index=False).agg(Consommation_kWh=("Energie_kWh","sum"), Puissance_moyenne_kW=("Puissance_kW","mean"), Puissance_max_kW=("Puissance_kW","max"))


def matrix_week_hour(h):
    x = h.copy(); x["Jour"] = x["Horodate"].dt.weekday.map(WEEKDAY_MAP); x["Heure"] = x["Horodate"].dt.hour
    return x.pivot_table(index="Heure", columns="Jour", values="Puissance_kW", aggfunc="mean").reindex(index=range(24), columns=WEEKDAYS)


def excel_export(raw, h, d, m, mat, meta):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        meta.to_excel(writer, index=False, sheet_name="Synthèse")
        raw.to_excel(writer, index=False, sheet_name="Données nettoyées")
        h.to_excel(writer, index=False, sheet_name="Profil horaire")
        d.to_excel(writer, index=False, sheet_name="Consommations journalières")
        m.to_excel(writer, index=False, sheet_name="Consommations mensuelles")
        mat.to_excel(writer, sheet_name="Moyenne heure-jour")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                width = min(max(max(len(str(c.value or "")) for c in cells)+2, 11), 34)
                ws.column_dimensions[cells[0].column_letter].width = width
    return out.getvalue()


def fr(v, n=1):
    return f"{v:,.{n}f}".replace(",", " ").replace(".", ",")


header()
st.markdown('<div class="card"><strong>Objectif :</strong> transformer un export Enedis en tableau de bord, profils horaires, consommations journalières et export Excel prêt à exploiter.</div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 1. Import")
    uploaded = st.file_uploader("Fichier Enedis", type=["csv","xlsx","xls"])

if uploaded is None:
    c1,c2,c3 = st.columns(3)
    for col,title,text in [(c1,"1. Importer","Déposez un export Enedis CSV ou Excel."),(c2,"2. Analyser","Les profils horaires et journaliers sont générés automatiquement."),(c3,"3. Exporter","Téléchargez un classeur Excel complet.")]:
        with col: st.markdown(f'<div class="card"><h3>{title}</h3><p>{text}</p></div>', unsafe_allow_html=True)
    st.info("Importez un fichier dans le panneau de gauche pour commencer.")
    st.stop()

try:
    base = read_file(uploaded.getvalue(), uploaded.name)
    step = detect_step(base); unit = detect_unit(base); base, unit_note = enrich(base, step, unit)
except Exception as e:
    st.error(f"Impossible de traiter le fichier : {e}"); st.stop()

with st.sidebar:
    st.markdown("---"); st.markdown("## 2. Période")
    mode = st.radio("Période d'analyse", ["Toutes les données", "Année", "Période personnalisée"])
    years = sorted(base["Horodate"].dt.year.unique().tolist())
    year = None; start = base["Horodate"].min().date(); end = base["Horodate"].max().date()
    if mode == "Année": year = st.selectbox("Année", years, index=len(years)-1)
    elif mode == "Période personnalisée":
        start = st.date_input("Date de début", start); end = st.date_input("Date de fin", end)
    st.markdown("---"); st.markdown("## 3. Hypothèse solaire")
    solar_start = st.slider("Début",0,23,8); solar_end = st.slider("Fin",1,24,18)

work = base.copy()
if mode == "Année": work = work[work["Horodate"].dt.year == year].copy()
elif mode == "Période personnalisée":
    work = work[(work["Horodate"] >= pd.Timestamp(start)) & (work["Horodate"] < pd.Timestamp(end)+pd.Timedelta(days=1))].copy()
if work.empty: st.warning("Aucune donnée sur la période choisie."); st.stop()

h = hourly(work); d = daily(work); m = monthly(work); mat = matrix_week_hour(h)
total = work["Energie_kWh"].sum(); avg_day = d["Consommation_kWh"].mean(); med_day = d["Consommation_kWh"].median(); peak = work["Puissance_kW"].max(); mean_kw = work["Puissance_kW"].mean()
solar = work.loc[(work["Horodate"].dt.hour >= solar_start) & (work["Horodate"].dt.hour < solar_end), "Energie_kWh"].sum(); solar_share = solar/total*100 if total else 0; load_factor = mean_kw/peak*100 if peak else 0
expected = round(pd.Timedelta(days=1)/step); by_day = work.assign(Date=work["Horodate"].dt.date).groupby("Date").size(); delta = round(pd.Timedelta(hours=1)/step); valid = [expected, expected-delta, expected+delta]; unusual = by_day[~by_day.isin(valid)]; duplicates = work["Horodate"].duplicated().sum()

c1,c2,c3 = st.columns([2,1,1]); c1.info(f"📅 Du **{work['Horodate'].min():%d/%m/%Y %H:%M}** au **{work['Horodate'].max():%d/%m/%Y %H:%M}**"); c2.info(f"⏱ **{int(step.total_seconds()/60)} min**"); c3.info(f"📐 **{unit}**"); st.caption(unit_note)

t1,t2,t3,t4,t5 = st.tabs(["📊 Tableau de bord","🕒 Profils horaires","📅 Journées","✅ Qualité","📥 Export"])

with t1:
    a,b,c,e,f = st.columns(5)
    a.metric("Consommation totale",f"{fr(total,0)} kWh"); b.metric("Moyenne journalière",f"{fr(avg_day)} kWh"); c.metric("Pic de puissance",f"{fr(peak)} kW"); e.metric(f"Part {solar_start}h–{solar_end}h",f"{fr(solar_share)} %"); f.metric("Facteur de charge",f"{fr(load_factor)} %")
    left,right = st.columns([1.45,1])
    with left:
        fig = px.bar(m,x="Mois_date",y="Consommation_kWh",title="Consommation mensuelle",labels={"Mois_date":"Mois","Consommation_kWh":"kWh"},color_discrete_sequence=[CMA_BLUE]); fig.update_layout(template="plotly_white",showlegend=False); st.plotly_chart(fig,use_container_width=True)
    with right:
        fig = go.Figure(go.Pie(labels=[f"Entre {solar_start}h et {solar_end}h","Hors plage solaire"],values=[solar,max(total-solar,0)],hole=.62,marker=dict(colors=[CMA_RED,"#EAF1F8"]),textinfo="label+percent")); fig.update_layout(title="Répartition de la consommation",template="plotly_white",showlegend=False); st.plotly_chart(fig,use_container_width=True)
    fig = px.line(h,x="Horodate",y="Puissance_kW",title="Évolution de la puissance moyenne horaire",labels={"Horodate":"Date et heure","Puissance_kW":"kW"},color_discrete_sequence=[CMA_RED]); fig.update_layout(template="plotly_white",hovermode="x unified"); st.plotly_chart(fig,use_container_width=True)

with t2:
    fig = go.Figure(go.Heatmap(z=mat.values,x=mat.columns,y=[f"{i:02d}:00" for i in mat.index],colorscale=[[0,"#5DAE74"],[.35,"#B8D96B"],[.58,"#F4E66B"],[.78,"#F7A65A"],[1,"#EF5A5A"]],colorbar=dict(title="kW"),hovertemplate="%{x}<br>%{y}<br>%{z:.2f} kW<extra></extra>")); fig.update_layout(title="Puissance moyenne selon le jour et l'heure",template="plotly_white",height=720); fig.update_yaxes(autorange="reversed"); st.plotly_chart(fig,use_container_width=True)
    long = mat.reset_index().melt(id_vars="Heure",var_name="Jour",value_name="Puissance_kW").dropna(); fig = px.line(long,x="Heure",y="Puissance_kW",color="Jour",markers=True,title="Profil de charge moyen sur 24 heures"); fig.update_layout(template="plotly_white",hovermode="x unified"); fig.update_xaxes(dtick=1); st.plotly_chart(fig,use_container_width=True)
    st.dataframe(mat.round(2),use_container_width=True,height=450)

with t3:
    a,b,c = st.columns(3); a.metric("Moyenne",f"{fr(avg_day)} kWh/j"); b.metric("Médiane",f"{fr(med_day)} kWh/j"); c.metric("Maximum",f"{fr(d['Consommation_kWh'].max())} kWh")
    fig = px.bar(d,x="Date",y="Consommation_kWh",color="Jour",category_orders={"Jour":WEEKDAYS},title="Consommation quotidienne",labels={"Consommation_kWh":"kWh"}); fig.update_layout(template="plotly_white",hovermode="x unified"); st.plotly_chart(fig,use_container_width=True)
    disp = d[["Date","Jour","Consommation_kWh","Puissance_moyenne_kW","Puissance_max_kW"]].copy(); disp["Date"] = disp["Date"].dt.strftime("%d/%m/%Y"); st.dataframe(disp,use_container_width=True,height=520)

with t4:
    a,b,c,e = st.columns(4); a.metric("Lignes exploitables",f"{len(work):,}".replace(","," ")); b.metric("Doublons",int(duplicates)); c.metric("Jours atypiques",len(unusual)); e.metric("Points théoriques/jour",expected)
    q = by_day.rename("Nombre_de_points").reset_index(); q.columns=["Date","Nombre_de_points"]; q["Statut"] = np.where(q["Nombre_de_points"].isin(valid),"Cohérent","À contrôler"); st.dataframe(q,use_container_width=True,height=450)
    if unusual.empty: st.success("Aucune journée anormale détectée en dehors des changements d'heure possibles.")
    else: st.warning("Certaines journées comportent un nombre de relevés inhabituel."); st.dataframe(unusual.rename("Nombre de relevés").to_frame(),use_container_width=True)

with t5:
    meta = pd.DataFrame({"Indicateur":["Fichier source","Début","Fin","Pas de temps","Unité source","Consommation totale (kWh)","Moyenne journalière (kWh)","Médiane journalière (kWh)","Pic de puissance (kW)",f"Part {solar_start}h-{solar_end}h (%)","Facteur de charge (%)","Doublons","Jours atypiques"],"Valeur":[uploaded.name,work["Horodate"].min(),work["Horodate"].max(),str(step),unit,total,avg_day,med_day,peak,solar_share,load_factor,duplicates,len(unusual)]})
    raw = work.copy(); raw["Horodate"] = raw["Horodate"].dt.strftime("%d/%m/%Y %H:%M:%S"); he = h.copy(); he["Horodate"] = he["Horodate"].dt.strftime("%d/%m/%Y %H:%M:%S"); de = d.copy(); de["Date"] = de["Date"].dt.strftime("%d/%m/%Y"); me = m.copy(); me["Mois_date"] = me["Mois_date"].dt.strftime("%m/%Y")
    xls = excel_export(raw,he,de,me,mat.round(3),meta)
    c1,c2 = st.columns(2)
    with c1: st.download_button("⬇️ Télécharger le classeur Excel complet",xls,"analyse_photovoltaique_cma.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    with c2: st.download_button("⬇️ Télécharger les consommations journalières",de.to_csv(index=False,sep=";").encode("utf-8-sig"),"consommations_journalieres.csv","text/csv",use_container_width=True)
    st.markdown("**Le classeur contient :** synthèse, données nettoyées, profil horaire, consommations journalières, consommations mensuelles et matrice heure × jour.")
