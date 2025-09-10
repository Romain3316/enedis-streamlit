import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Traitement Enedis", layout="centered")

st.title("🔧 Traitement des relevés de charge Enedis")

uploaded_file = st.file_uploader("📁 Déposez votre fichier ici (.csv ou .xlsx)", type=["csv", "xlsx"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8', low_memory=False)
        else:
            df = pd.read_excel(uploaded_file)

        df = df[['Unité', 'Horodate', 'Valeur']]
        df = df[df['Unité'].str.upper().isin(['W', 'KW'])]
        df['Horodate'] = pd.to_datetime(df['Horodate'], errors='coerce')
        df['Valeur'] = pd.to_numeric(df['Valeur'], errors='coerce')
        df = df.dropna()

        df['Tranche_Horaire'] = (df['Horodate'] + pd.Timedelta(minutes=60)).dt.floor('H')
        df_hourly = df.groupby(['Unité', 'Tranche_Horaire'])['Valeur'].mean().reset_index()
        df_hourly['Date'] = df_hourly['Tranche_Horaire'].dt.date
        df_hourly['Heure'] = df_hourly['Tranche_Horaire'].dt.time
        df_hourly = df_hourly[['Unité', 'Date', 'Heure', 'Valeur']]
        df_hourly.rename(columns={'Valeur': 'Moyenne de consommation'}, inplace=True)

        df_hourly['Tranche_Horaire'] = pd.to_datetime(df_hourly['Date'].astype(str) + ' ' + df_hourly['Heure'].astype(str))
        df_2024 = df_hourly[(df_hourly['Tranche_Horaire'] >= '2024-01-01') & (df_hourly['Tranche_Horaire'] < '2025-01-01')]
        df_final = df_2024[['Unité', 'Date', 'Heure', 'Moyenne de consommation']]

        output = BytesIO()
        df_final.to_csv(output, index=False)
        output.seek(0)

        st.success("✅ Fichier traité avec succès.")
        st.download_button("📥 Télécharger le CSV", output, file_name="consommation_2024.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Une erreur est survenue : {e}")
