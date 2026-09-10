"""Streamlit integration: restore before widgets, retain hidden settings."""

import streamlit as st
from dossiers import decode_dossier, encode_dossier, source_from_record
from dossier_schema import SETTINGS


def initialize_dossier():
    pending = st.session_state.pop("_pending_dossier", None)
    if pending:
        if pending["version"] == 2:
            for key in SETTINGS:
                st.session_state.pop(key, None)
            st.session_state.pop("address_candidates", None)
            st.session_state.pop("_resolved_location", None)
            st.session_state["_solar_snapshot"] = pending.get("solar_snapshot")
            st.session_state["_source_files"] = pending["files"]
            st.session_state["_upload_revision"] = st.session_state.get("_upload_revision", 0) + 1
        st.session_state.update(pending["fields"])
        st.session_state["_dossier_message"] = (
            "Dossier complet restauré : fichiers, paramètres et annotations."
            if pending["version"] == 2 else
            "Ancien brouillon restauré : les fichiers source et les réglages restent à vérifier."
        )
    # Explicitly detach persisted keys from widget cleanup when a control is hidden.
    for key in SETTINGS:
        if key in st.session_state:
            st.session_state[key] = st.session_state[key]


def _load_dossier():
    source = st.session_state.get("dossier_upload")
    if source is None: return
    try:
        st.session_state["_pending_dossier"] = decode_dossier(source.getvalue())
    except ValueError as exc:
        st.session_state["_dossier_error"] = str(exc)


def render_dossier_loader():
    st.markdown("### Reprendre un dossier")
    st.file_uploader("Dossier complet ou ancien brouillon (.json)", type=["json"], key="dossier_upload")
    st.button("Charger le dossier", on_click=_load_dossier, key="load_dossier", use_container_width=True)
    if "_dossier_error" in st.session_state: st.error(st.session_state.pop("_dossier_error"))
    if "_dossier_message" in st.session_state: st.success(st.session_state.pop("_dossier_message"))


def _update_source(role, key):
    uploaded = st.session_state.get(key)
    files = dict(st.session_state.get("_source_files", {}))
    if uploaded is None: files.pop(role, None)
    else: files[role] = {"name": uploaded.name, "content": uploaded.getvalue()}
    st.session_state["_source_files"] = files


def _remove_source(role):
    files = dict(st.session_state.get("_source_files", {}))
    files.pop(role, None)
    st.session_state["_source_files"] = files
    st.session_state["_upload_revision"] = st.session_state.get("_upload_revision", 0) + 1


def source_uploader(role, label, extensions, help_text):
    key = f"source_{role}_{st.session_state.get('_upload_revision', 0)}"
    st.file_uploader(label, type=extensions, help=help_text, key=key,
                     on_change=_update_source, args=(role, key))
    record = st.session_state.get("_source_files", {}).get(role)
    if record:
        st.caption(f"Fichier utilisé : {record['name']}")
        st.button("Retirer ce fichier", key=f"remove_{role}", on_click=_remove_source, args=(role,))
    return source_from_record(record)


def render_dossier_download():
    with st.sidebar:
        st.markdown("### Sauvegarder le dossier")
        fields = {key: st.session_state[key] for key in SETTINGS if key in st.session_state}
        # The selected geocoded location becomes portable manual coordinates.
        location = st.session_state.get("_resolved_location")
        if location and not fields.get("manual_coordinates"):
            fields.update(manual_coordinates=True, manual_latitude=location['latitude'], manual_longitude=location['longitude'])
        try:
            data = encode_dossier(fields, st.session_state.get("_source_files", {}), st.session_state.get("_solar_snapshot"))
        except ValueError as exc:
            st.error(f"Sauvegarde impossible : {exc}")
            return
        name = ''.join(c if c.isalnum() or c in '-_' else '_' for c in fields.get('company_name', 'entreprise')) or 'entreprise'
        st.download_button("Sauvegarder le dossier complet (.json)", data=data,
            file_name=f"dossier_analyse_energetique_{name}.json", mime="application/json", key="save_complete_dossier", use_container_width=True)
        st.caption("Conserve les fichiers de consommation et PMA, les paramètres, les notes et le profil PVGIS déjà obtenu. Téléchargez à nouveau le dossier après vos modifications.")
