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
    # Lecture CSV ou Excel
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", encoding="utf-8", low_memory=False)
    else:
        df = pd.read_excel(uploaded_file)

    # 2. Garder seulement colonnes utiles
    df = df[["Unité", "Horodate", "Valeur"]]

    # 3. Supprimer "VAR"
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]

    # 4. Conversion date
    df["Horodate"] = pd.to_datetime(df["Horodate"], dayfirst=True, errors="coerce")

    # 5. Agrégation horaire → moyenne
    df = df.set_index("Horodate")
    df = df.resample("1H").mean(numeric_only=True)
    df = df.dropna().reset_index()

    # Années disponibles
    annees_dispo = sorted(df["Horodate"].dt.year.unique().tolist())

    # Widgets Streamlit
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

    # Bouton export
    if st.button("🚀 Lancer le traitement"):
        # Filtrage période
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee]

        elif choix_periode == "Période personnalisée":
            debut = pd.to_datetime(date_debut)
            fin = pd.to_datetime(date_fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

        # Gestion des jours 23h/25h
        if mode_horaire == "Forcer 24h/jour":
            full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
            df = df.set_index("Horodate").reindex(full_range)
            df.index.name = "Horodate"
            df["Valeur"] = df["Valeur"].interpolate(method="linear")
            df = df.reset_index()

        # Vérification des trous
        trous = []
        full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
        missing = full_range.difference(df["Horodate"])
        if not missing.empty:
            trous = [d.strftime("%d/%m/%Y %H:%M:%S") for d in missing]

        # Format final
        df["Date"] = df["Horodate"].dt.date
        df["Heure"] = df["Horodate"].dt.time
        df = df.rename(columns={"Valeur": "Moyenne_Conso"})
        df_final = df[["Unité", "Date", "Heure", "Moyenne_Conso"]]

        # Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20))

        # Message trous
        if trous:
            st.warning(f"⚠️ Données manquantes (exemple) : {', '.join(trous[:5])}")
        else:
            st.success("✅ Pas de données manquantes")

        # Export
        if format_export == "CSV":
            csv = df_final.to_csv(index=False, sep=";").encode("utf-8")
            st.download_button("⬇️ Télécharger en CSV", csv, "donnees_enedis.csv", "text/csv")

        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("⬇️ Télécharger en Excel", output.getvalue(), "donnees_enedis.xlsx", "application/vnd.ms-excel")
