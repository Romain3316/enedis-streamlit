import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Traitement des données Enedis", layout="wide")

st.title("📊 Traitement des données Enedis")

# 1. Import fichier
uploaded_file = st.file_uploader(
    "Choisissez un fichier Enedis (Excel ou CSV)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    usecols = ["Unité", "Horodate", "Valeur", "Pas"]

    # ✅ Lecture CSV ou Excel
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", usecols=lambda c: c in usecols, dtype="string")
    else:
        df = pd.read_excel(uploaded_file, usecols=lambda c: c in usecols, dtype="string")

    # 2. Conversion datetime
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)

    # ✅ Supprimer la toute première ligne (23h55 → 00h00 de la veille)
    df = df.iloc[1:].reset_index(drop=True)

    # 3. Nettoyage → garder uniquement W et kW + dates valides
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df = df.dropna(subset=["Horodate", "Valeur"]).sort_values("Horodate").reset_index(drop=True)

    # 4. Détection du pas de temps
    if len(df) > 1:
        diffs = df["Horodate"].diff().dropna()
        pas_detecte = diffs.mode()[0]  # le pas le plus fréquent
        # format lisible (ex : 5min ou 1h)
        if pas_detecte.seconds % 3600 == 0:
            pas_affiche = f"{pas_detecte.seconds // 3600}h"
        else:
            pas_affiche = f"{pas_detecte.seconds // 60} min"
    else:
        pas_affiche = "inconnu"

    st.info(
        f"📅 Données disponibles : du **{df['Horodate'].min().strftime('%d/%m/%Y %H:%M')}** "
        f"au **{df['Horodate'].max().strftime('%d/%m/%Y %H:%M')}**\n\n"
        f"⏱ Pas de temps détecté : **{pas_affiche}**"
    )

    # 5. Années disponibles
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())

    # 6. Widgets Streamlit
    choix_periode = st.selectbox(
        "📅 Choisissez la période à exporter :",
        ["Toutes les données"] + [str(a) for a in annees_dispo] + ["Période personnalisée"]
    )

    mode_horaire = st.radio(
        "⏱ Gestion des jours à 23h / 25h :",
        ["Heures réelles (23h / 25h)", "Forcer 24h/jour"]
    )

    format_export = st.radio("📂 Format d'export :", ["CSV", "Excel"])

    if choix_periode == "Période personnalisée":
        col1, col2 = st.columns(2)
        with col1:
            date_debut = st.date_input("Date de début", value=df["Horodate"].min().date())
        with col2:
            date_fin = st.date_input("Date de fin", value=df["Horodate"].max().date())

    if st.button("🚀 Lancer le traitement"):

        # 7. Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee]
        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

        # 8. Agrégation
        if mode_horaire == "Heures réelles (23h / 25h)":
            df["Horodate_hour"] = df["Horodate"].dt.floor("H") + pd.Timedelta(hours=1)
            df = df.groupby("Horodate_hour", as_index=False)["Valeur"].mean()
            df = df.rename(columns={"Horodate_hour": "Horodate"})
        else:
            full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
            df = df.set
