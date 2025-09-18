# 9. Diagnostic des heures par jour
heures_par_jour = df.groupby(df["Horodate"].dt.date).size()

# ✅ Détection spécifique aux changements d'heure (23h ou 25h)
changements_heure = heures_par_jour[heures_par_jour.isin([23, 25])]

# ✅ Détection des jours vraiment incomplets (hors changements d'heure)
jours_incomplets = heures_par_jour[~heures_par_jour.isin([23, 24, 25])]

st.subheader("📊 Diagnostic des heures par jour")

if changements_heure.empty and jours_incomplets.empty:
    st.success("Toutes les journées comptent 24 heures (aucun changement d'heure détecté).")
else:
    if not changements_heure.empty:
        st.warning("⚠️ Changements d'heure détectés (23h ou 25h) :")
        st.dataframe(changements_heure)

    if not jours_incomplets.empty:
        st.error("❌ Jours incomplets détectés (moins de 23h ou plus de 25h) :")
        st.dataframe(jours_incomplets)
