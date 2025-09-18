import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt

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
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)

    # 3. Nettoyage
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df = df.dropna(subset=["Horodate", "Valeur"]).reset_index(drop=True)

    # ⚡ Supprimer la première ligne (correspond au 23h55 → 00h00 de la veille)
    df = df.iloc[1:].reset_index(drop=True)

    # ⚡ Trouver le premier jour complet
    heures_par_jour = df.groupby(df["Horodate"].dt.date).size()
    jours_valides = heures_par_jour[heures_par_jour == 24]

    if not jours_valides.empty:
        premier_jour_valide = jours_valides.index.min()
        df = df[df["Horodate"].dt.date >= premier_jour_valide].reset_index(drop=True)
    else:
        st.warning("⚠️ Aucun jour complet détecté → conservation de toutes les données après la 1ʳᵉ ligne.")

    # 4. Vérification des bornes
    debut_brut, fin_brut = df["Horodate"].min(), df["Horodate"].max()
    if pd.notna(debut_brut) and pd.notna(fin_brut):
        st.info(f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
                f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**")
    else:
        st.error("⚠️ Impossible de déterminer les bornes des données après filtrage.")
        st.stop()

    # 5. Pas de temps
    pas_detecte = df["Horodate"].diff().min()
    if pd.notna(pas_detecte):
        st.info(f"⏱ Pas de temps détecté : {pas_detecte.components.minutes} min")
    else:
        st.warning("⏱ Impossible de détecter le pas de temps.")

    # 6. Widgets Streamlit
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())
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

        # 7. Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee]
        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

        # 8. Mode horaire
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

        # 9. Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]
        df_final = df[["Date", "Heure", "Moyenne_Conso"]]

        # 10. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # 11. Courbe complète
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

        # 12. Heatmap jour × heure
        st.subheader("🔥 Heatmap de la consommation (jour vs heure)")
        df["JourSemaine"] = df["Horodate"].dt.dayofweek
        df["HeureJour"] = df["Horodate"].dt.hour
        heatmap_data = df.pivot_table(
            index="JourSemaine", columns="HeureJour", values="Valeur", aggfunc="mean"
        )

        plt.figure(figsize=(14, 5))
        sns.heatmap(heatmap_data, cmap="RdYlGn_r", cbar_kws={"label": "Conso moyenne"})
        plt.xlabel("Heure de la journée")
        plt.ylabel("Jour de la semaine (0=Lundi)")
        st.pyplot(plt)

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
