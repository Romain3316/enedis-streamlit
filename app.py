import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px

st.title("📊 Traitement des données Enedis")

# ==========================
# 1. Import fichier
# ==========================
uploaded_file = st.file_uploader(
    "Choisissez un fichier Enedis (Excel ou CSV)",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    usecols = ["Unité", "Horodate", "Valeur"]

    # ✅ Lecture CSV ou Excel
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", usecols=usecols, dtype={"Unité": "string"})
    else:
        df = pd.read_excel(uploaded_file, usecols=usecols, dtype={"Unité": "string"})

    # 2. Conversion datetime
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["Horodate", "Valeur"])
    df = df.sort_values("Horodate")

    # 3. Détection du pas de temps
    pas = df["Horodate"].diff().mode()[0]
    st.info(f"⏱ Pas de temps détecté : {pas}")

    # 4. Décalage des horodatages (chaque valeur est la conso jusqu’à l’horodate)
    df["Horodate_corrige"] = df["Horodate"] - pas

    # ⚠️ Supprimer la première ligne après décalage (elle correspondrait à 23h-00h mais est incomplète)
    df = df.iloc[1:].reset_index(drop=True)

    # 5. Vérification des bornes après correction
    debut_brut, fin_brut = df["Horodate_corrige"].min(), df["Horodate"].max()
    st.info(f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
            f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**")

    # 6. Années disponibles
    annees_dispo = sorted(df["Horodate_corrige"].dt.year.unique().tolist())

    # ==========================
    # Widgets Streamlit
    # ==========================
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
            date_debut = st.date_input("Date de début", value=df["Horodate_corrige"].min().date())
        with col2:
            date_fin = st.date_input("Date de fin", value=df["Horodate_corrige"].max().date())

    # ==========================
    # Bouton traitement
    # ==========================
    if st.button("🚀 Lancer le traitement"):

        # 7. Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate_corrige"].dt.year == annee]
        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate_corrige"] >= debut) & (df["Horodate_corrige"] <= fin)]

        # 8. Agrégation selon le mode choisi
        if mode_horaire == "Heures réelles (23h / 25h)":
            df["Horodate_hour"] = df["Horodate_corrige"].dt.floor("H") + pd.Timedelta(hours=1)
            df_grouped = df.groupby("Horodate_hour", as_index=False)["Valeur"].mean()
            df_grouped = df_grouped.rename(columns={"Horodate_hour": "Horodate"})
        else:
            full_range = pd.date_range(df["Horodate_corrige"].min(), df["Horodate_corrige"].max(), freq="1H")
            df_grouped = df.set_index("Horodate_corrige").resample("1H").mean(numeric_only=True).reindex(full_range)
            df_grouped.index.name = "Horodate"
            df_grouped["Valeur"] = df_grouped["Valeur"].interpolate(method="linear")
            df_grouped = df_grouped.reset_index()

        # 9. Diagnostic des heures
        heures_par_jour = df_grouped.groupby(df_grouped["Horodate"].dt.date).size()
        jours_suspects = heures_par_jour[(heures_par_jour < 23) | (heures_par_jour > 25)]

        st.subheader("📊 Changements d'heure détectés")
        if jours_suspects.empty:
            st.success("✅ Aucun changement d'heure détecté sur la période.")
        else:
            st.warning("⚠️ Changements d'heure détectés :")
            st.dataframe(jours_suspects)

        # 10. Format final
        df_grouped["Date"] = df_grouped["Horodate"].dt.strftime("%d/%m/%Y")
        df_grouped["Heure"] = df_grouped["Horodate"].dt.strftime("%H:%M:%S")
        df_grouped["Moyenne_Conso"] = df_grouped["Valeur"]

        df_final = df_grouped[["Date", "Heure", "Moyenne_Conso"]]

        # 11. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # 12. Courbe sur l’ensemble des données
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
            template="plotly_white",
            hovermode="x unified"
        )
        st.plotly_chart(fig_full, use_container_width=True)

        # 13. Export
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
