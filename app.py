import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px

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
        df = pd.read_csv(uploaded_file, sep=";", usecols=usecols, dtype={"Unité": "string", "Pas": "string"})
    else:
        df = pd.read_excel(uploaded_file, usecols=usecols, dtype={"Unité": "string", "Pas": "string"})

    # 2. Conversion datetime en JJ/MM/AAAA
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)

    # ⚡ Décalage automatique selon la colonne "Pas"
    def get_offset(pas):
        if pas == "PT5M":
            return pd.Timedelta(minutes=5)
        elif pas == "PT10M":
            return pd.Timedelta(minutes=10)
        elif pas == "PT15M":
            return pd.Timedelta(minutes=15)
        elif pas == "PT30M":
            return pd.Timedelta(minutes=30)
        elif pas in ["PT60M", "PT1H"]:
            return pd.Timedelta(hours=1)
        else:
            return pd.Timedelta(0)

    df["Offset"] = df["Pas"].apply(get_offset)

    # ✅ Décalage au milieu du pas
    df["Horodate"] = df["Horodate"] - (df["Offset"] / 2)

    # 🚨 Supprimer la première ligne (consommation partielle de la veille)
    df = df.iloc[1:].reset_index(drop=True)

    # 3. Nettoyage → garder uniquement W et kW
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df = df.dropna(subset=["Horodate", "Valeur"])

    # ⚡ Correction du problème de démarrage
    debut_brut, fin_brut = df["Horodate"].min(), df["Horodate"].max()
    st.info(f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
            f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**")

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
            df["Horodate_hour"] = df["Horodate"].dt.floor("H")
            df = df.groupby("Horodate_hour", as_index=False)["Valeur"].mean()
            df = df.rename(columns={"Horodate_hour": "Horodate"})
        else:
            full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
            df = df.set_index("Horodate").resample("1H").mean(numeric_only=True).reindex(full_range)
            df.index.name = "Horodate"
            df["Valeur"] = df["Valeur"].interpolate(method="linear")
            df = df.reset_index()

        # 9. Diagnostic : uniquement changements d'heure
        heures_par_jour = df.groupby(df["Horodate"].dt.date).size()
        changements_heure = heures_par_jour[(heures_par_jour == 23) | (heures_par_jour == 25)]

        st.subheader("⏳ Changements d'heure détectés")
        if changements_heure.empty:
            st.success("✅ Aucun changement d'heure détecté sur la période.")
        else:
            st.warning("⚠️ Changements d'heure détectés (23h ou 25h) :")
            st.dataframe(changements_heure)

        # 10. Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]

        df_final = df[["Date", "Heure", "Moyenne_Conso"]]

        # 11. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # 12. Graphique global
        df_plot = df_final.copy()
        df_plot["Datetime"] = pd.to_datetime(df_plot["Date"] + " " + df_plot["Heure"], dayfirst=True)

        fig_full = px.line(
            df_plot,
            x="Datetime",
            y="Moyenne_Conso",
            title="📈 Évolution de la consommation (ensemble des données)",
        )
        fig_full.update_traces(line=dict(width=2))
        fig_full.update_layout(
            xaxis_title="Date et heure",
            yaxis_title="Consommation moyenne",
            template="plotly_dark",
            hovermode="x unified"
        )
        st.plotly_chart(fig_full, use_container_width=True)

        # 13. Heatmap jour/heure
        st.subheader("🔥 Heatmap de la consommation (jour vs heure)")
        jours_fr = {0: "Lundi", 1: "Mardi", 2: "Mercredi", 3: "Jeudi", 4: "Vendredi", 5: "Samedi", 6: "Dimanche"}
        df["JourSemaine"] = df["Horodate"].dt.dayofweek.map(jours_fr)
        df["HeureJour"] = df["Horodate"].dt.hour

        heatmap_data = df.pivot_table(
            index="HeureJour", columns="JourSemaine", values="Moyenne_Conso", aggfunc="mean"
        )

        fig_heatmap = px.imshow(
            heatmap_data,
            labels=dict(x="Jour de semaine", y="Heure de la journée", color="Conso moyenne"),
            aspect="auto",
            color_continuous_scale="RdYlGn_r"
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

        # 14. Export
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
