"""Espace expérimental de lecture et de validation des factures."""
import re
import streamlit as st
from dossier_schema import SETTINGS
from invoice_reader import read_invoice
from invoice_metadata import FIELDS, extract_metadata

PRICE_FIELDS = {
    "HP": ("hp_electricity_price", "Heures pleines"),
    "HC": ("hc_electricity_price", "Heures creuses"),
    "HPE": ("hp_summer_electricity_price", "HP été"),
    "HCE": ("hc_summer_electricity_price", "HC été"),
    "HPH": ("hp_winter_electricity_price", "HP hiver"),
    "HCH": ("hc_winter_electricity_price", "HC hiver"),
}


def clear_invoice():
    for key in list(st.session_state):
        if key.startswith("_invoice_") and key != "_invoice_upload":
            del st.session_state[key]


def reset_invoice():
    clear_invoice()
    st.session_state.pop("_invoice_upload", None)


def render_invoice_tests():
    st.header("Factures — Tests")
    st.caption("Déposez une facture, vérifiez la lecture, puis choisissez les valeurs à appliquer au dossier courant.")
    st.info("Version expérimentale : reconnaissance des lignes Endesa et EDF HP/HC lisibles. "
            "Les autres formats restent consultables dans le texte extrait, sans préremplissage garanti.")
    source = st.file_uploader("Facture PDF ou photo JPG/PNG (10 Mo, PDF de 6 pages maximum)", type=["pdf", "jpg", "jpeg", "png"],
                              key="_invoice_upload", on_change=clear_invoice)
    st.button("Effacer la facture de test", key="clear_invoice", on_click=reset_invoice)
    st.caption("Une facture à la fois. Lecture locale au serveur, sans service d’IA externe. "
               "Le document et le texte extrait ne sont pas intégrés au dossier JSON sauvegardé.")
    if st.button("Lire la facture", key="read_invoice", disabled=source is None):
        clear_invoice()
        try:
            with st.spinner("Lecture du PDF et reconnaissance de texte si nécessaire…"):
                result = read_invoice(source.getvalue())
            st.session_state["_invoice_result"] = result
        except ValueError as exc:
            st.error(str(exc))
    result = st.session_state.get("_invoice_result")
    if not result:
        if any(st.session_state.get(k) for k in FIELDS if k != "company_name"):
            render_invoice_information("")
        return
    st.caption(f"{result['pages']} page(s) lue(s) · OCR : " +
               (", ".join(map(str, result["ocr_pages"])) if result["ocr_pages"] else "non nécessaire"))
    with st.expander("Vérifier le texte extrait", expanded=not result["rows"]):
        st.text(result["text"])
    if result["rows"]:
        st.dataframe([{"Plage": r["series"], "Du": r["start"], "Au": r["end"],
                       "Consommation (kWh)": r["kwh"], "Prix lu (€/kWh)": r["price"],
                       "Ligne source à vérifier": r["evidence"]} for r in result["rows"]],
                     use_container_width=True, hide_index=True)
    render_invoice_information(result["text"])
    candidates = result["candidates"]
    if not candidates:
        st.warning("Aucun tarif fiable reconnu par plage. Vous pouvez saisir les prix ci-dessous après lecture de la facture. Ne déduisez pas HP/HC de l’ordre des nombres dans le texte OCR.")
    st.subheader("Proposer un préremplissage")
    st.warning("Vérifiez le périmètre des prix : Endesa distingue la fourniture de l’acheminement ; "
               "EDF peut déjà inclure l’acheminement dans ses prix. Ne l’ajoutez pas deux fois. "
               "Les taxes ne sont pas ajoutées automatiquement. "
               "Ne reportez pas le coût moyen TTC de la facture comme prix du kWh.")
    st.caption("Aucune conversion kW → kVA. Abonnement, frais fixes et horaires HP/HC restent inchangés. "
               "Les plages non reconnues conservent leurs paramètres actuels, qui peuvent être des valeurs par défaut.")
    models = ["HP / HC", "HP / HC hiver-été"]
    default_model = 1 if any(k in candidates for k in ("HPE", "HCE", "HPH", "HCH")) else 0
    model = st.selectbox("Structure tarifaire à vérifier", models, index=default_model, key="_invoice_model")
    fields = {k: v for k, v in PRICE_FIELDS.items() if (len(k) == 2) == (model == "HP / HC")}
    with st.form("invoice_review"):
        selected = st.multiselect("Plages à remplacer", list(fields), default=[], key="_invoice_selected")
        values = {}
        for series, (key, label) in fields.items():
            current = st.session_state.get(key, SETTINGS[key]["default"])
            values[key] = st.number_input(
                    f"{label} — prix à appliquer (€/kWh hors TVA)", min_value=0.0, max_value=5.0,
                    value=float(candidates[series]["price"]) if series in candidates else None, format="%.5f", step=0.001,
                    key="_invoice_price_" + series,
                    help=f"Paramètre actuel : {current:.5f} €/kWh. Vérifiez la ligne source et le périmètre du prix.",
                )
            if series not in candidates:
                st.caption(f"{label} : non extrait. Laissez vide pour conserver {current:.5f} €/kWh.")
        st.caption(f"L’application passera au tarif « {model} ». Vérifiez aussi les prix "
                   "manquants et les plages horaires dans l’analyse avant d’utiliser ses résultats.")
        verified = st.checkbox("J’ai vérifié les prix hors TVA, les unités et les paramètres saisonniers conservés.",
                               key="_invoice_verified")
        submitted = st.form_submit_button("Appliquer les valeurs sélectionnées au dossier")
    if submitted:
        if not verified or not selected or any(values[fields[s][0]] is None for s in selected):
            st.error("Renseignez chaque plage sélectionnée et confirmez la vérification avant d’appliquer.")
        else:
            for series in selected:
                key = PRICE_FIELDS[series][0]
                st.session_state[key] = values[key]
            st.session_state["electricity_tariff_type"] = model
            st.session_state.pop("_prepared_export", None)
            st.success("Prix sélectionnés appliqués. Revenez dans l’analyse pour vérifier les horaires et "
                       "les autres paramètres, puis sauvegardez votre dossier.")


def render_invoice_information(text):
    st.subheader("Informations de la facture — à vérifier")
    proposed = extract_metadata(text)
    st.caption("Les champs vides n’ont pas été reconnus avec certitude. Corrigez la lecture et sélectionnez uniquement les informations à enregistrer. Les valeurs validées seront sauvegardées avec le dossier.")
    identity = [key for key in ("company_name", "invoice_address") if key in proposed]
    if identity:
        st.info("Nom et adresse repérés : les champs détectés sont préremplis et présélectionnés ci-dessous. Vérifiez-les puis enregistrez-les pour les reporter dans le dossier.")
        for key in identity:
            st.text(FIELDS[key] + " : " + proposed[key]["value"])
        st.caption("Pour reporter également l’adresse dans l’analyse, cochez la confirmation « site étudié ».")
    if proposed:
        with st.expander("Sources des informations proposées"):
            st.dataframe([{"Champ": FIELDS[k], "Proposition": v["value"], "Texte source": v["evidence"]}
                          for k, v in proposed.items()], hide_index=True, use_container_width=True)
    with st.form("invoice_information"):
        selected_labels = st.multiselect("Informations à enregistrer", list(FIELDS.values()),
                                        default=[FIELDS[key] for key in identity],
                                        key="_invoice_info_selected")
        selected = [key for key, label in FIELDS.items() if label in selected_labels]
        values = {key: st.text_input(label, value=proposed.get(key, {}).get("value", "") if text else st.session_state.get(key, ""),
                                    key="_invoice_info_" + key,
                                    help="Valeur actuelle du dossier : " + (st.session_state.get(key, "") or "non renseignée"))
                  for key, label in FIELDS.items()}
        site_address = st.checkbox("J’ai vérifié que cette adresse est celle du site étudié : l’utiliser aussi comme adresse de l’entreprise.", key="_invoice_info_site")
        st.caption("L’adresse de facturation peut différer du site alimenté. Le PDL/PRM doit correspondre au compteur étudié. La puissance reste descriptive : aucune conversion kW/kVA ni modification des calculs.")
        verified = st.checkbox("J’ai vérifié ces informations sur la facture et leur correspondance avec ce dossier.", key="_invoice_info_verified")
        submitted = st.form_submit_button("Enregistrer les informations validées")
    if submitted:
        if not verified or not selected or any(not values[k].strip() for k in selected):
            st.error("Sélectionnez des informations renseignées et confirmez leur vérification.")
        elif "invoice_prm" in selected and not re.fullmatch(r"\d{14}", values["invoice_prm"].strip()):
            st.error("Le PDL/PRM doit comporter exactement 14 chiffres.")
        elif site_address and "invoice_address" not in selected:
            st.error("Sélectionnez l’adresse parmi les informations à enregistrer.")
        else:
            for key in selected:
                st.session_state[key] = values[key].strip()
            if site_address:
                st.session_state["company_address"] = values["invoice_address"].strip()
                st.session_state.pop("address_candidates", None)
                st.session_state.pop("selected_label", None)
                st.session_state["manual_coordinates"] = False
                st.session_state.pop("_resolved_location", None)
                st.session_state.pop("_solar_snapshot", None)
            st.session_state.pop("_prepared_export", None)
            st.success("Informations enregistrées dans le dossier courant. Sauvegardez le dossier pour les conserver.")
            if site_address:
                st.info("Recherchez et validez la nouvelle adresse dans l’analyse photovoltaïque avant de recalculer.")
