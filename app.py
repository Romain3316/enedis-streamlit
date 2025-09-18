import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt

# ==============================
# 🎨 Personnalisation CMA
# ==============================
CMA_COLOR = "#981C31"
CMA_BG = "#F9F9F9"

st.set_page_config(
    page_title="CMA Nouvelle-Aquitaine - Données Enedis",
    page_icon="📊",
    layout="wide",
)

# CSS custom
st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: {CMA_BG};
        color: black;
    }}
    h1, h2, h3, h4, h5, h6, label, .stRadio label, .stSelectbox label {{
        color: black !important;
        font-weight: bold;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# Logo
st.image("logo-cma-na.png", width=250)

st.title("📊 Traitement des données Enedis")

# 1. Import fichier
uploaded_file = st.file_uploader(
    "Choisissez un fichier Enedis (Excel ou CSV)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    # ✅ Lecture rapide avec seulement les colonnes utiles
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", usecols=["Unité", "Horodate", "Valeur"])
    else:
        df = pd.read_excel(uploaded_file, usecols=["Unité", "Horodate", "Valeur"])

    # 2. Nettoyage
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df["Horodate"] = pd.to_datetime(df["Horodate"], format="%d/%m/%Y %H:%M", errors="coerce")
    df = df.dropna(subset=["Horodate"])

    # 3. Agrégation horaire brute
    df = df.set_index("Horodate").resample("1H").mean(numeric_only=True).reset_index()

    # 4. Années disponibles
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())

    # 5. Widgets Streamlit
    choix_periode = st.selectbox(
        "📅 Choisissez la période à exporter :",
        ["Toutes les données"] + [str(a) for a in annees_dispo] + ["Période personnalisée"]
    )

    mode_horaire = st.radio(
        "⏱ Gestion des jours à 23h / 25h :",
        ["Heures réelles (23h / 25h)", "Forcer 24h/jour"]
    )

    if choix_periode == "Période personnalisée":
        col1, col2 = st.columns(2)
        with col1:
            date_debut = st.date_input("Date de début", value=df["Horodate"].min().date())
        with col2:
            date_fin = st.date_input("Date de fin", value=df["Horodate"].max().date())

    # 🚀 Bouton de traitement
    if st.button("🚀 Lancer le traitement"):

        # === Filtrage selon le choix ===
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee]

        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

        # 6. Gestion des trous si "forcer 24h/jour"
        if mode_horaire == "Forcer 24h/jour":
            full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
            df = df.set_index("Horodate").reindex(full_range)
            df.index.name = "Horodate"
            df["Valeur"] = df["Valeur"].interpolate(method="linear")
            df = df.reset_index()

        # 7. Colonnes finales
        df["Date"] = df["Horodate"].dt.date
        df["Heure"] = df["Horodate"].dt.time
        df["Moyenne_Conso"] = df["Valeur"]

        # 📋 Aperçu tableau
        st.subheader("📋 Aperçu des 20 premières données traitées")
        st.dataframe(df[["Date", "Heure", "Moyenne_Conso"]].head(20))

        # 📈 Courbe complète
        st.subheader("📈 Évolution de la consommation")
        fig = px.line(
            df, 
            x="Horodate", 
            y="Moyenne_Conso", 
            title="Évolution de la consommation (toutes les données)",
            template="plotly_dark"
        )
        fig.update_traces(line=dict(color=CMA_COLOR, width=2))
        st.plotly_chart(fig, use_container_width=True)

        # 🔥 Heatmap hebdomadaire
        st.subheader("🔥 Heatmap hebdomadaire")
        df_heatmap = df.dropna(subset=["Moyenne_Conso"]).copy()
        df_heatmap["Jour_semaine"] = df_heatmap["Horodate"].dt.dayofweek
        df_heatmap["Heure"] = df_heatmap["Horodate"].dt.hour

        pivot = df_heatmap.pivot_table(
            values="Moyenne_Conso", 
            index="Jour_semaine", 
            columns="Heure", 
            aggfunc="mean"
        )

        plt.figure(figsize=(14,6))
        plt.imshow(pivot, aspect="auto", cmap="RdYlGn_r")
        plt.colorbar(label="Consommation (Moyenne)")
        plt.xlabel("Heure")
        plt.ylabel("Jour de semaine (0=Lundi ... 6=Dimanche)")
        plt.title("Heatmap de la consommation par heure et jour de semaine", fontsize=14, fontweight="bold")
        st.pyplot(plt.gcf())
