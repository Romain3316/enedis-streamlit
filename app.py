import pandas as pd
import streamlit as st
from io import BytesIO

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

    # 2. Conversion datetime en JJ/MM/AAAA
    df["Horodate"] = pd.to_datetime(df["Horodate"], errors="coerce", dayfirst=True)

    # 3. Nettoyage → garder uniquement W et kW
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]
    df = df.dropna(subset=["Horodate", "Valeur"])

    # 4. Vérification des bornes
    debut_brut, fin_brut = df["Horodate"].min(), df["Horodate"].max()
    st.info(f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
            f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**")

    # 5. Agrégation horaire → moyenne
    df_resampled = df.set_index("Horodate").resample("1H").mean(numeric_only=True).reset_index()

    # ⚡ Correction : démarrer au premier horaire complet
    if not df_resampled.empty:
        debut = df_resampled["Horodate"].min()
        df_resampled = df_resampled[df_resampled["Horodate"] >= debut.ceil("1H")]

    df = df_resampled

    # 6. Années disponibles
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())

    # 7. Widgets Streamlit
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
        # 8. Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee]
        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

        # 9. Gestion des jours 23h/25h
        if mode_horaire == "Forcer 24h/jour":
            if not df.empty:
                full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
                df = df.set_index("Horodate").reindex(full_range)
                df.index.name = "Horodate"
                df["Valeur"] = df["Valeur"].interpolate(method="linear")
                df = df.reset_index()

        # 10. Vérification des anomalies
        anomalies = []
        if not df.empty:
            full_range = pd.date_range(
                df["Horodate"].min(),
                df["Horodate"].max(),
                freq="1H",
                tz="Europe/Paris"
            ).tz_convert(None)

            missing = full_range.difference(df["Horodate"])
            missing = missing[(missing > df["Horodate"].min()) & (missing < df["Horodate"].max())]
            if not missing.empty:
                anomalies.extend([f"Heure manquante (été): {d.strftime('%d/%m/%Y %H:%M:%S')}" for d in missing])

            df["Jour"] = df["Horodate"].dt.date
            df["Heure_str"] = df["Horodate"].dt.strftime("%H:%M")
            doublons_horaires = df[df.duplicated(subset=["Jour", "Heure_str"], keep=False)]
            if not doublons_horaires.empty:
                for jour in doublons_horaires["Jour"].unique():
                    anomalies.append(f"Heure doublée (hiver): {jour} 02:00–03:00")

        # 11. Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]

        if "Unité" in df.columns:
            df_final = df[["Unité", "Date", "Heure", "Moyenne_Conso"]]
        else:
            df_final = df[["Date", "Heure", "Moyenne_Conso"]]

        # 12. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # 13. Message anomalies
        if anomalies:
            st.warning("⚠️ Anomalies détectées :\n" + "\n".join(anomalies))
        else:
            st.success("✅ Pas de données manquantes ni doublées")

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
