import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="Traitement des données Enedis", layout="wide")
st.title("📊 Traitement des données Enedis")

# 1. Import du fichier
uploaded_file = st.file_uploader(
    "Choisissez un fichier Enedis (Excel ou CSV)",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
    usecols = ["Unité", "Horodate", "Valeur"]

    try:
        # Lecture CSV ou Excel
        if uploaded_file.name.lower().endswith(".csv"):
            df_source = pd.read_csv(
                uploaded_file,
                sep=";",
                usecols=usecols,
                dtype={"Unité": "string"},
            )
        else:
            df_source = pd.read_excel(
                uploaded_file,
                usecols=usecols,
                dtype={"Unité": "string"},
            )
    except (ValueError, KeyError) as exc:
        st.error(
            "Le fichier doit contenir les colonnes « Unité », « Horodate » "
            f"et « Valeur ». Détail : {exc}"
        )
        st.stop()
    except Exception as exc:
        st.error(f"Impossible de lire le fichier : {exc}")
        st.stop()

    # 2. Nettoyage et conversion
    df_source["Horodate"] = pd.to_datetime(
        df_source["Horodate"],
        errors="coerce",
        dayfirst=True,
    )

    # Accepte aussi les nombres écrits avec une virgule décimale
    df_source["Valeur"] = pd.to_numeric(
        df_source["Valeur"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )

    nb_lignes_initiales = len(df_source)
    df_source = df_source.dropna(subset=["Horodate", "Valeur"]).copy()
    nb_lignes_ignorees = nb_lignes_initiales - len(df_source)

    if df_source.empty:
        st.error("Aucune donnée exploitable n'a été trouvée dans le fichier.")
        st.stop()

    if nb_lignes_ignorees:
        st.warning(
            f"⚠️ {nb_lignes_ignorees} ligne(s) invalide(s) ont été ignorée(s)."
        )

    # 3. Tri chronologique
    df_source = df_source.sort_values("Horodate").reset_index(drop=True)

    # 4. Détection robuste du pas de temps
    ecarts = df_source["Horodate"].diff()
    ecarts_positifs = ecarts[ecarts > pd.Timedelta(0)]

    if ecarts_positifs.empty:
        st.error("Impossible de détecter le pas de temps.")
        st.stop()

    # Le mode est plus fiable que le minimum en présence d'anomalies ponctuelles
    modes = ecarts_positifs.mode()
    pas = modes.iloc[0] if not modes.empty else ecarts_positifs.median()

    # IMPORTANT :
    # On conserve les horodatages Enedis tels qu'ils figurent dans le fichier.
    # L'ancienne version supprimait la première ligne et décalait les heures,
    # ce qui provoquait la disparition de 01:00 et le décalage des valeurs.
    debut_brut = df_source["Horodate"].min()
    fin_brut = df_source["Horodate"].max()

    st.info(
        f"📅 Données disponibles : du **{debut_brut.strftime('%d/%m/%Y %H:%M')}** "
        f"au **{fin_brut.strftime('%d/%m/%Y %H:%M')}**"
    )
    st.info(f"⏱ Pas de temps détecté : **{pas.total_seconds() / 60:g} min**")

    # 5. Paramètres de traitement
    annees_dispo = sorted(df_source["Horodate"].dt.year.unique().tolist())

    choix_periode = st.selectbox(
        "📅 Choisissez la période à exporter :",
        ["Toutes les données"]
        + [str(annee) for annee in annees_dispo]
        + ["Période personnalisée"],
    )

    mode_horaire = st.radio(
        "⏱ Gestion des jours à 23 h / 25 h :",
        ["Heures réelles (23 h / 25 h)", "Forcer 24 h/jour"],
    )

    format_export = st.radio("📂 Format d'export :", ["CSV", "Excel"])

    if choix_periode == "Période personnalisée":
        col1, col2 = st.columns(2)

        with col1:
            date_debut = st.date_input(
                "Date de début",
                value=df_source["Horodate"].min().date(),
            )

        with col2:
            date_fin = st.date_input(
                "Date de fin",
                value=df_source["Horodate"].max().date(),
            )

    # 6. Traitement
    if st.button("🚀 Lancer le traitement"):
        df = df_source.copy()

        # Filtrage de la période sur les horodatages d'origine
        if choix_periode not in ["Toutes les données", "Période personnalisée"]:
            annee = int(choix_periode)
            df = df[df["Horodate"].dt.year == annee].copy()

        elif choix_periode == "Période personnalisée":
            debut = pd.Timestamp(date_debut)
            fin_exclue = pd.Timestamp(date_fin) + pd.Timedelta(days=1)

            df = df[
                (df["Horodate"] >= debut)
                & (df["Horodate"] < fin_exclue)
            ].copy()

        if df.empty:
            st.error("Aucune donnée ne correspond à la période choisie.")
            st.stop()

        # 7. Agrégation horaire
        if mode_horaire == "Heures réelles (23 h / 25 h)":
            # Une donnée horodatée à 01:00 reste à 01:00.
            # Si plusieurs relevés existent dans la même heure, on en calcule la moyenne.
            df["Horodate_heure"] = df["Horodate"].dt.floor("h")

            df = (
                df.groupby("Horodate_heure", as_index=False, sort=True)["Valeur"]
                .mean()
                .rename(columns={"Horodate_heure": "Horodate"})
            )

        else:
            debut_plage = df["Horodate"].min().floor("h")
            fin_plage = df["Horodate"].max().floor("h")

            plage_complete = pd.date_range(
                start=debut_plage,
                end=fin_plage,
                freq="1h",
            )

            df = (
                df.set_index("Horodate")
                .resample("1h")["Valeur"]
                .mean()
                .reindex(plage_complete)
                .rename_axis("Horodate")
                .to_frame()
            )

            df["Valeur"] = df["Valeur"].interpolate(
                method="time",
                limit_direction="both",
            )

            df = df.reset_index()

        # 8. Diagnostic des jours incomplets et changements d'heure
        heures_par_jour = df.groupby(df["Horodate"].dt.date).size()

        jours_23_25 = heures_par_jour[heures_par_jour.isin([23, 25])]
        jours_incomplets = heures_par_jour[
            ~heures_par_jour.isin([23, 24, 25])
        ]

        st.subheader("📊 Contrôle du nombre d'heures par jour")

        if jours_23_25.empty and jours_incomplets.empty:
            st.success("✅ Tous les jours complets comportent 24 heures.")
        else:
            if not jours_23_25.empty:
                st.info("🕒 Journées de changement d'heure détectées :")
                st.dataframe(
                    jours_23_25.rename("Nombre d'heures").to_frame(),
                    use_container_width=True,
                )

            if not jours_incomplets.empty:
                st.warning(
                    "⚠️ Journées comportant moins de 23 heures ou plus de 25 heures :"
                )
                st.dataframe(
                    jours_incomplets.rename("Nombre d'heures").to_frame(),
                    use_container_width=True,
                )

        # 9. Format final
        df["Date"] = df["Horodate"].dt.strftime("%d/%m/%Y")
        df["Heure"] = df["Horodate"].dt.strftime("%H:%M:%S")
        df["Moyenne_Conso"] = df["Valeur"]

        df_final = df[["Date", "Heure", "Moyenne_Conso"]].copy()

        # 10. Aperçu
        st.subheader("📋 Aperçu des données traitées")
        st.dataframe(df_final.head(20), use_container_width=True)

        # 11. Graphique
        df_plot = df[["Horodate", "Valeur"]].copy()

        fig_full = px.line(
            df_plot,
            x="Horodate",
            y="Valeur",
            title="📈 Évolution de la consommation",
        )
        fig_full.update_traces(line=dict(width=2))
        fig_full.update_layout(
            xaxis_title="Date et heure",
            yaxis_title="Consommation moyenne",
            template="plotly_dark",
            hovermode="x unified",
        )

        st.plotly_chart(fig_full, use_container_width=True)

        # 12. Export
        if format_export == "CSV":
            csv = df_final.to_csv(index=False, sep=";").encode("utf-8-sig")

            st.download_button(
                "⬇️ Télécharger en CSV",
                data=csv,
                file_name="donnees_enedis.csv",
                mime="text/csv",
            )

        else:
            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(
                    writer,
                    index=False,
                    sheet_name="Données traitées",
                )

            st.download_button(
                "⬇️ Télécharger en Excel",
                data=output.getvalue(),
                file_name="donnees_enedis.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
