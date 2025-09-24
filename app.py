import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="Traitement des données Enedis", layout="wide")
st.title("📊 Traitement des données Enedis")

# 1. Import fichier
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
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce")
    df = df.dropna(subset=["Horodate", "Valeur"])

    # 3. Tri chrono
    df = df.sort_values("Horodate").reset_index(drop=True)

    # 4. Détection du pas de temps
    pas = df["Horodate"].diff().min()
    if pd.isna(pas):
        st.error("Impossible de détecter le pas de temps.")
        st.stop()

    # Supprimer la première ligne car elle représente la conso de la veille
    df = df.iloc[1:].reset_index(drop=True)

    # Décaler les horodatages restants
    df["Horodate_corrige"] = df["Horodate"] - pas

    # 5. Vérification des bornes
    debut_brut, fin_brut = df["Horodate_corrige"].min(), df["Horodate_corrige"].max()
    st.info(f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
            f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**")
    st.info(f"⏱ Pas de temps détecté : {pas.components.hours*60 + pas.components.minutes} min")

    # 6. Années disponibles
    annees_dispo = sorted(df["Horodate_corrige"].dt.year.unique().tolist())
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
            date_debut = st.date_input("Date de début", value=df["Horodate_corrige"].min().date())
        with col2:
            date_fin = st.date_input("Date de fin", value=df["Horodate_corrige"].max().date())

    # Bouton traitement
    if st.button("🚀 Lancer le traitement"):

        # 7. Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate_corrige"].dt.year == annee]
        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate_corrige"] >= debut) & (df["Horodate_corrige"] <= fin)]

        # 8. Agrégation horaire
        if mode_horaire == "Heures réelles (23h / 25h)":
            df["Horodate_hour"] = df["Horodate_corrige"].dt.floor("H") + pd.Timedelta(hours=1)
            df = df.groupby("Horodate_hour", as_index=False)["Valeur"].mean()
            df = df.rename(columns={"Horodate_hour": "Horodate"})
        else:
            full_range = pd.date_range(df["Horodate_corrige"].min(), df["Horodate_corrige"].max(), freq="1H")
            df = df.set_index("Horodate_corrige").resample("1H").mean(numeric_only=True).reindex(full_range)
            df.index.name = "Horodate"
            df["Valeur"] = df["Valeur"].interpolate(method="linear")
            df = df.reset_index()

        # 9. Diagnostic changements d’heure
        heures_par_jour = df.groupby(df["Horodate"].dt.date).size()
        jours_suspects = heures_par_jour[(heures_par_jour < 23) | (heures_par_jour > 25)]

        st.subheader("📊 Changements d'heure détectés")
        if jours_suspects.empty:
            st.success("✅ Aucun changement d'heure détecté sur la période.")
        else:
            st.warning("⚠️ Changements d'heure détectés :")
            st.dataframe(jours_suspects)

        # 10. Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]

        df_final = df[["Date", "Heure", "Moyenne_Conso"]]

        # 11. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # 12. Graph complet
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
