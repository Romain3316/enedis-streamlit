import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px

# ==============================
# 🎨 Personnalisation CMA
# ==============================
CMA_COLOR = "#9B1C31"
CMA_BG = "#F9F9F9"

st.set_page_config(
    page_title="CMA Nouvelle-Aquitaine - Données Enedis",
    page_icon="⚡",
    layout="wide",
)

# CSS custom
st.markdown(
    f"""
    <style>
        .reportview-container {{
            background-color: {CMA_BG};
        }}
        .stButton button {{
            background-color: {CMA_COLOR};
            color: white;
            border-radius: 8px;
            padding: 0.6em 1.2em;
            font-weight: bold;
        }}
        .stButton button:hover {{
            background-color: #7C1527;
            color: white;
        }}
        h1, h2, h3, h4 {{
            color: {CMA_COLOR};
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==============================
# 🖼️ Logo CMA
# ==============================
st.image("logo_cma.png", width=250)
st.title("📊 Analyse des données Enedis - CMA Nouvelle-Aquitaine")

# ==============================
# 📂 Import fichier
# ==============================
uploaded_file = st.file_uploader(
    "Choisissez un fichier Enedis (Excel ou CSV)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    usecols = ["Unité", "Horodate", "Valeur"]

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", usecols=usecols, dtype={"Unité": "string"})
    else:
        df = pd.read_excel(uploaded_file, usecols=usecols, dtype={"Unité": "string"})

    # Conversion datetime
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)

    # Filtre W / kW
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df = df.dropna(subset=["Horodate", "Valeur"])

    # Dates disponibles
    debut_brut, fin_brut = df["Horodate"].min(), df["Horodate"].max()
    st.info(f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
            f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**")

    # Années disponibles
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())

    # Widgets
    choix_periode = st.selectbox(
        "📅 Choisissez la période à exporter :",
        ["Toutes les données"] + [str(a) for a in annees_dispo] + ["Période personnalisée"]
    )

    mode_horaire = st.radio(
        "⏱ Gestion des jours à 23h / 25h :",
        ["Heures réelles (23h / 25h)", "Forcer 24h/jour"]
    )

    format_export = st.radio("📂 Format d'export :", ["CSV", "Excel"])

    # Période personnalisée
    if choix_periode == "Période personnalisée":
        col1, col2 = st.columns(2)
        with col1:
            date_debut = st.date_input("Date de début", value=df["Horodate"].min().date())
        with col2:
            date_fin = st.date_input("Date de fin", value=df["Horodate"].max().date())

    # Bouton traitement
    if st.button("🚀 Lancer le traitement"):

        # Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee]
        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

        # Agrégation
        if mode_horaire == "Heures réelles (23h / 25h)":
            df["Horodate_hour"] = df["Horodate"].dt.floor("H") + pd.Timedelta(hours=1)
            df = df.groupby("Horodate_hour", as_index=False)["Valeur"].mean()
            df = df.rename(columns={"Horodate_hour": "Horodate"})
        else:
            full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
            df = df.set_index("Horodate").resample("1H").mean(numeric_only=True).reindex(full_range)
            df.index.name = "Horodate"
            df["Valeur"] = df["Valeur"].interpolate(method="linear")
            df = df.reset_index()

        # Diagnostic des heures par jour
        heures_par_jour = df.groupby(df["Horodate"].dt.date).size()
        jours_suspects = heures_par_jour[heures_par_jour != 24]

        st.subheader("📊 Diagnostic des heures par jour")
        if jours_suspects.empty:
            st.success("Toutes les journées comptent 24 heures (mode forcé ou période sans changement d'heure).")
        else:
            st.warning("⚠️ Jours avec un nombre d'heures différent de 24 détectés :")
            st.dataframe(jours_suspects)

        # Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]

        df_final = df[["Date", "Heure", "Moyenne_Conso"]]

        # ==============================
        # 📋 Tableau des 20 premières données
        # ==============================
        st.subheader("📋 Aperçu des 20 premières données traitées")
        st.dataframe(df_final.head(20))

        # ==============================
        # 📈 Graphique complet
        # ==============================
        st.subheader("📈 Évolution de la consommation (ensemble des données)")
        df_plot = df_final.copy()
        df_plot["Datetime"] = pd.to_datetime(df_plot["Date"] + " " + df_plot["Heure"], dayfirst=True)

        fig_full = px.line(
            df_plot,
            x="Datetime",
            y="Moyenne_Conso",
            title="📈 Évolution de la consommation",
        )
        fig_full.update_traces(line=dict(width=2, color=CMA_COLOR))
        fig_full.update_layout(
            xaxis_title="Date et heure",
            yaxis_title="Consommation moyenne",
            template="simple_white",
            hovermode="x unified"
        )
        st.plotly_chart(fig_full, use_container_width=True)

        # ==============================
        # 🔥 Heatmap hebdo (traduction FR sans locale)
        # ==============================
        st.subheader("🔥 Profil hebdomadaire de consommation")
        jours_fr = {
            "Monday": "Lundi",
            "Tuesday": "Mardi",
            "Wednesday": "Mercredi",
            "Thursday": "Jeudi",
            "Friday": "Vendredi",
            "Saturday": "Samedi",
            "Sunday": "Dimanche"
        }
        df_heatmap = df.copy()
        df_heatmap["Jour_semaine"] = df_heatmap["Horodate"].dt.day_name().map(jours_fr)
        df_heatmap["Heure"] = df_heatmap["Horodate"].dt.strftime("%H:00")

        pivot = df_heatmap.pivot_table(
            index="Heure",
            columns="Jour_semaine",
            values="Valeur",
            aggfunc="mean"
        )

        ordre_jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
        pivot = pivot[[j for j in ordre_jours if j in pivot.columns]]

        fig_heatmap = px.imshow(
            pivot,
            labels=dict(x="Jour de la semaine", y="Heure", color="Conso moyenne"),
            aspect="auto",
            color_continuous_scale="RdYlGn_r",
        )
        fig_heatmap.update_layout(
            title="🔥 Profil hebdomadaire",
            xaxis_title="Jour de la semaine",
            yaxis_title="Heure de la journée"
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

        # Export
        if format_export == "CSV":
            csv = df_final.to_csv(index=False, sep=";").encode("utf-8")
            st.download_button("⬇️ Télécharger en CSV", csv, "donnees_enedis.csv", "text/csv")
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False)
            st.download_button(
                "⬇️ Télécharger en Excel",
                output.getvalue(),
                "donnees_enedis.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

