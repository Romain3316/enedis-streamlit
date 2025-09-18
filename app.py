import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

# ==============================
# 🎨 Personnalisation CMA
# ==============================
CMA_COLOR = "#9B1C31"   # Rouge CMA
CMA_BG = "#F9F9F9"      # Fond clair
TEXT_COLOR = "#000000"  # Noir

st.set_page_config(
    page_title="CMA Nouvelle-Aquitaine - Données Enedis",
    page_icon="⚡",
    layout="wide",
)

# CSS custom
st.markdown(f"""
    <style>
        .stApp {{
            background-color: {CMA_BG};
            color: {TEXT_COLOR};
        }}
        .stButton>button {{
            background-color: {CMA_COLOR};
            color: white;
            font-weight: bold;
            border-radius: 8px;
        }}
        .stRadio label, .stSelectbox label, .stDateInput label {{
            color: {TEXT_COLOR};
            font-weight: bold;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: {TEXT_COLOR};
        }}
    </style>
""", unsafe_allow_html=True)

# Logo CMA
st.image("logo-cma-na.png", width=250)

# ==============================
# 📊 Application
# ==============================
st.title("⚡ Traitement des données Enedis - CMA Nouvelle-Aquitaine")

uploaded_file = st.file_uploader(
    "Choisissez un fichier Enedis (Excel ou CSV)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    # ✅ Lecture uniquement des colonnes utiles
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", usecols=["Unité", "Horodate", "Valeur"])
    else:
        df = pd.read_excel(uploaded_file, usecols=["Unité", "Horodate", "Valeur"])

    # ✅ Forçage parsing JJ/MM/AAAA
    df["Horodate"] = pd.to_datetime(df["Horodate"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Horodate"])
    df = df.sort_values("Horodate").drop_duplicates("Horodate")

    # ✅ Suppression des "VAR", on garde W / kW
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]

    # ✅ On coupe avant le 12/06/2023 si besoin
    df = df[df["Horodate"] >= pd.to_datetime("2023-06-12")]

    # Agrégation horaire
    df = df.set_index("Horodate").resample("1H").mean(numeric_only=True).reset_index()

    # Colonnes utiles
    df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
    df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
    df["Moyenne_Conso"] = df["Valeur"]

    # ==============================
    # 📅 Choix période
    # ==============================
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())
    choix_periode = st.selectbox(
        "📅 Choisissez la période à exporter :",
        ["Toutes les données"] + [str(a) for a in annees_dispo] + ["Période personnalisée"]
    )

    if choix_periode == "Période personnalisée":
        col1, col2 = st.columns(2)
        with col1:
            date_debut = st.date_input("Date de début", value=df["Horodate"].min().date())
        with col2:
            date_fin = st.date_input("Date de fin", value=df["Horodate"].max().date())

    format_export = st.radio("📂 Format d'export :", ["CSV", "Excel"])

    # ==============================
    # 🚀 Traitement
    # ==============================
    if st.button("🚀 Lancer le traitement"):
        data = df.copy()

        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            data = data[data["Horodate"].dt.year == annee]

        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            data = data[(data["Horodate"] >= debut) & (data["Horodate"] <= fin)]

        # ==============================
        # ⚠️ Diagnostic jours incomplets
        # ==============================
        diag = data.groupby(data["Horodate"].dt.date).size()
        jours_incomplets = diag[diag != 24]

        if not jours_incomplets.empty:
            st.warning("⚠️ Jours avec un nombre d'heures différent de 24 détectés :")
            st.dataframe(jours_incomplets)
        else:
            st.success("✅ Tous les jours ont bien 24 relevés")

        # ==============================
        # 📋 Tableau aperçu
        # ==============================
        st.subheader("📋 Aperçu des 20 premières données traitées")
        st.dataframe(data[["Date", "Heure", "Moyenne_Conso"]].head(20))

        # ==============================
        # 📈 Graphique stylisé (TOUTES données sélectionnées)
        # ==============================
        st.subheader("📈 Évolution de la consommation")
        fig = px.line(
            data,
            x="Horodate", y="Moyenne_Conso",
            title="Évolution de la consommation",
            line_shape="spline"
        )
        fig.update_traces(line=dict(width=2, color=CMA_COLOR))
        fig.update_layout(
            title_font=dict(size=18, color=TEXT_COLOR),
            xaxis_title="Date",
            yaxis_title="Consommation (Moyenne)",
            font=dict(color=TEXT_COLOR)
        )
        st.plotly_chart(fig, use_container_width=True)

        # ==============================
        # 🔥 Heatmap hebdo (Seaborn)
        # ==============================
        st.subheader("🔥 Heatmap hebdomadaire")
        df_heatmap = data.copy()
        df_heatmap["Jour_semaine"] = df_heatmap["Horodate"].dt.dayofweek  # 0 = Lundi
        df_heatmap["Heure"] = df_heatmap["Horodate"].dt.hour

        pivot = df_heatmap.pivot_table(
            index="Heure", columns="Jour_semaine", values="Moyenne_Conso", aggfunc="mean"
        )

        jours_labels = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"]

        plt.figure(figsize=(12,6))
        sns.heatmap(pivot, cmap="RdYlGn_r", annot=False, xticklabels=jours_labels)
        plt.title("Heatmap de la consommation par heure et jour de semaine", fontsize=14, color=TEXT_COLOR)
        plt.xlabel("Jour de la semaine", color=TEXT_COLOR)
        plt.ylabel("Heure", color=TEXT_COLOR)
        st.pyplot(plt)

        # ==============================
        # 📂 Export
        # ==============================
        df_final = data[["Date", "Heure", "Moyenne_Conso"]]

        if format_export == "CSV":
            csv = df_final.to_csv(index=False, sep=";").encode("utf-8")
            st.download_button("⬇️ Télécharger en CSV", csv, "donnees_enedis.csv", "text/csv")

        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("⬇️ Télécharger en Excel", output.getvalue(),
                               "donnees_enedis.xlsx",
                               "application/vnd.ms-excel")
