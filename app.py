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
    usecols = ["Unité", "Horodate", "Valeur"]

    # ✅ Lecture CSV ou Excel
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", usecols=usecols, dtype={"Unité": "string"})
    else:
        df = pd.read_excel(uploaded_file, usecols=usecols, dtype={"Unité": "string"})

    # 2. Conversion datetime
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)

    # 3. Nettoyage → garder uniquement W et kW
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df = df.dropna(subset=["Horodate", "Valeur"])

    # ⚠️ Correction : ne commencer qu’au 12/06/2023
    df = df[df["Horodate"] >= pd.Timestamp("2023-06-12 00:00:00")]

    # 4. Vérification des bornes
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

        # 8. Agrégation selon le mode choisi
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

        # 9. Diagnostic des heures par jour
        heures_par_jour = df.groupby(df["Horodate"].dt.date).size()

        jours_23_25 = heures_par_jour[(heures_par_jour == 23) | (heures_par_jour == 25)]
        jours_incomplets = heures_par_jour[(heures_par_jour < 23) | (heures_par_jour > 25)]

        st.subheader("📊 Diagnostic des heures par jour")
        if jours_23_25.empty and jours_incomplets.empty:
            st.success("✅ Toutes les journées comptent exactement 24 heures.")
        else:
            if not jours_23_25.empty:
                st.warning("⚠️ Jours avec 23h ou 25h détectés (changements d’heure) :")
                st.dataframe(jours_23_25)
            if not jours_incomplets.empty:
                st.error("❌ Jours incomplets détectés (moins de 23h ou plus de 25h) :")
                st.dataframe(jours_incomplets)

        # 10. Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]

        df_final = df[["Date", "Heure", "Moyenne_Conso"]]

        # 11. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # 12. Courbe de prévisualisation (20 lignes)
        preview = df_final.head(20).copy()
        preview["Datetime"] = pd.to_datetime(preview["Date"] + " " + preview["Heure"], dayfirst=True)

        fig_preview = px.line(
            preview,
            x="Datetime",
            y="Moyenne_Conso",
            markers=True,
            title="⚡ Évolution de la consommation (aperçu sur 20 lignes)",
        )
        fig_preview.update_traces(line=dict(width=3), fill="tozeroy")
        fig_preview.update_layout(
            xaxis_title="Date et heure",
            yaxis_title="Consommation moyenne",
            template="plotly_dark",
            hovermode="x unified"
        )
        st.plotly_chart(fig_preview, use_container_width=True)

        # 13. Courbe stylisée sur l’ensemble des données
        df_plot = df_final.copy()
        df_plot["Datetime"] = pd.to_datetime(df_plot["Date"] + " " + df_plot["Heure"], dayfirst=True)

        # Lissage 24h
        df_plot["Conso_Smooth"] = df_plot["Moyenne_Conso"].rolling(window=24, min_periods=1).mean()

        fig_full = px.area(
            df_plot,
            x="Datetime",
            y="Conso_Smooth",
            title="📈 Évolution de la consommation (ensemble des données, lissé 24h)",
        )

        fig_full.update_traces(line=dict(width=2, color="crimson"), fill="tozeroy", opacity=0.7)
        fig_full.update_layout(
            xaxis_title="Date et heure",
            yaxis_title="Consommation moyenne (lissée)",
            template="plotly_dark",
            hovermode="x unified"
        )
        st.plotly_chart(fig_full, use_container_width=True)

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
