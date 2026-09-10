from datetime import datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock

import pandas as pd
import pytest
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest

import dossier_ui
from dossiers import decode_dossier, encode_dossier

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text() + '''
st.session_state["_test_result"] = dict(
    energy_pdf=globals().get("energy_pdf_bytes"), pv_pdf=globals().get("pdf_report_bytes"), excel=globals().get("excel_bytes"),
    variable=tariff_variable_cost_eur, fixed=tariff_fixed_cost_eur,
    total=total_kwh, tariff=tariff_summary_df.to_dict("records"))
'''


def curve_file():
    dates = pd.date_range("2025-01-01 01:00", periods=168, freq="h")
    data = pd.DataFrame({"Horodate": dates.strftime("%d/%m/%Y %H:%M"), "Unité": "kW", "Valeur": 5.0})
    return {"name": "synthetic.csv", "content": data.to_csv(index=False, sep=";").encode()}


def pma_file():
    return {"name": "synthetic-pma.csv", "content": b"Horodate;Grandeur physique;Valeur;Unite\n2025-01-01 12:00:00;PMA;25000;VA\n2025-01-02 12:00:00;PMA;24000;VA\n"}


def launch(fields=None, files=None):
    at = AppTest.from_string(APP, default_timeout=90)
    at.session_state["_workspace_page"] = "Rapport"
    at.session_state["_pending_dossier"] = decode_dossier(encode_dossier(fields or {}, files or {"curve": curve_file()}))
    return at.run()


def assert_ok(at):
    assert not at.exception, [e.message for e in at.exception]
    assert not at.error, [e.value for e in at.error]


@pytest.mark.parametrize("tariff,expected", [("Tarif unique", 0.31 * 840), ("HP / HC", (16*0.4+8*0.2)*5*7), ("HP / HC hiver-été", (16*0.5+8*0.1)*5*7)])
def test_energy_is_offline_and_uses_selected_tariff(monkeypatch, tariff, expected):
    get = Mock(side_effect=AssertionError("Le parcours énergie ne doit pas appeler le réseau"))
    monkeypatch.setattr("requests.get", get)
    at = launch({"electricity_tariff_type": tariff, "unique_electricity_price": 0.31,
        "hp_electricity_price": 0.4, "hc_electricity_price": 0.2,
        "hp_winter_electricity_price": 0.5, "hc_winter_electricity_price": 0.1,
        "annual_subscription_eur": 365.25})
    assert_ok(at)
    get.assert_not_called()
    result = at.session_state["_test_result"]
    assert result["variable"] == pytest.approx(expected)
    assert result["fixed"] == pytest.approx(7)
    assert result["pv_pdf"] is None
    assert "Pré-diagnostic énergétique" in PdfReader(BytesIO(result["energy_pdf"])).pages[0].extract_text()
    assert "Opportunité photovoltaïque" not in [t.label for t in at.tabs]
    assert "pv_peak_kwp" not in [w.key for w in at.number_input]
    workbook = pd.ExcelFile(BytesIO(result["excel"]))
    assert "Projection financière" not in workbook.sheet_names
    summary = pd.read_excel(workbook, sheet_name="Synthèse").set_index("Indicateur")["Valeur"]
    assert "Puissance photovoltaïque étudiée (kWc)" not in summary.index
    assert summary["Total sur la période (€ HT)"] == pytest.approx(expected + 7)


def solar_response():
    response = Mock()
    dates = pd.date_range("2025-01-01", periods=240, freq="h")
    response.json.return_value = {"inputs": {}, "outputs": {"hourly": [
        {"time": d.strftime("%Y%m%d:%H%M"), "P": 3000 if 8 <= d.hour < 17 else 0, "G(i)": 400 if 8 <= d.hour < 17 else 0}
        for d in dates]}}
    return response


def test_switch_restore_and_pdf_annotations(monkeypatch, tmp_path):
    get = Mock(return_value=solar_response())
    monkeypatch.setattr("requests.get", get)
    captured = []
    original_encode = dossier_ui.encode_dossier
    def capture(*args, **kwargs):
        raw = original_encode(*args, **kwargs)
        captured.append(raw)
        return raw
    monkeypatch.setattr(dossier_ui, "encode_dossier", capture)
    fields = {"analysis_mode": "Opportunité photovoltaïque", "manual_coordinates": True,
              "pv_peak_kwp": 17.5, "hc_range_count": 2, "hc2_start": "12:00:00", "hc2_end": "14:00:00",
              "company_name": "Atelier de démonstration", "company_address": "Adresse fictive pour vérification du rapport",
              "note_context": "CONTEXTE_TEST : activité artisanale.", "note_profile": "PROFIL_TEST : consommation régulière.",
              "note_power": "PUISSANCE_TEST : vérifier les démarrages.", "note_tariff": "TARIF_TEST : contrat à confirmer.",
              "note_investigate": "QUESTIONS_TEST : relever les équipements.", "note_pv_observations": "PV_TEST : étude toiture nécessaire.",
              "note_recommendations": "ACTIONS_TEST : programmer une visite.", "note_general": "GENERAL_TEST : données fictives.",
              "report_status": "À compléter après rendez-vous"}
    fields["note_context"] += "\n\n" + ("Le conseiller vérifie les horaires, les équipements et les usages avec le dirigeant. " * 14)
    at = launch(fields, {"curve": curve_file(), "pma": pma_file()})
    assert_ok(at)
    first = at.session_state["_test_result"]
    assert first["pv_pdf"]
    text = "\n".join(p.extract_text() for p in PdfReader(BytesIO(first["pv_pdf"])).pages)
    for left, note, right in [
        ("1. Synthèse", "CONTEXTE_TEST", "2. Comprendre"),
        ("2. Comprendre", "PROFIL_TEST", "3. Potentiel"),
        ("2. Comprendre", "PUISSANCE_TEST", "3. Potentiel"),
        ("3. Potentiel", "PV_TEST", "4. Analyse tarifaire"),
        ("4. Analyse tarifaire", "TARIF_TEST", "5. Étude financière"),
        ("10. Préparation", "QUESTIONS_TEST", "11. Limites"),
    ]:
        assert text.index(left) < text.index(note) < text.index(right)
    assert "À compléter après rendez-vous" in text
    # Keep the generated PDFs as temporary QA output when explicitly requested.
    import os
    if os.environ.get("QA_PDF_DIR"):
        out = Path(os.environ["QA_PDF_DIR"]); out.mkdir(parents=True, exist_ok=True)
        (out / "energie.pdf").write_bytes(first["energy_pdf"])
        (out / "photovoltaique.pdf").write_bytes(first["pv_pdf"])
    at.checkbox(key="_include_pv").uncheck().run()
    assert_ok(at)
    at.radio(key="hc_range_count").set_value(1).run()
    assert_ok(at)
    saved = captured[-1]
    restored = decode_dossier(saved)
    assert restored["fields"]["pv_peak_kwp"] == 17.5
    assert restored["fields"]["note_pv_observations"] == fields["note_pv_observations"]
    assert restored["files"]["pma"] == pma_file()
    assert restored["solar_snapshot"]
    get.reset_mock(side_effect=True)
    get.side_effect = AssertionError("Le profil sauvegardé doit être réutilisé")
    fresh = AppTest.from_string(APP, default_timeout=90)
    fresh.session_state["_workspace_page"] = "Rapport"
    fresh.session_state["_pending_dossier"] = restored
    fresh.run()
    assert_ok(fresh)
    fresh.checkbox(key="_include_pv").check().run()
    assert_ok(fresh)
    fresh.radio(key="hc_range_count").set_value(2).run()
    assert_ok(fresh)
    assert fresh.time_input(key="hc2_start").value.hour == 12
    assert fresh.number_input(key="pv_peak_kwp").value == 17.5
    assert fresh.session_state["_test_result"]["variable"] == first["variable"]
    get.assert_not_called()
    # Preview must use the same already-generated report and keep notes intact.
    fresh.checkbox(key="show_energy_pdf").check().run()
    assert_ok(fresh)
    fresh.checkbox(key="show_pv_pdf").check().run()
    assert_ok(fresh)


def test_loading_another_dossier_clears_old_sources_and_notes():
    at = launch({"note_context": "Note du premier dossier", "company_name": "Premier"},
                {"curve": curve_file(), "pma": pma_file()})
    assert_ok(at)
    at.session_state["_workspace_page"] = "Rapport"
    at.session_state["_pending_dossier"] = decode_dossier(encode_dossier({"company_name": "Second"}, {"curve": curve_file()}))
    at.run()
    assert_ok(at)
    assert at.text_area(key="note_context").value == ""
    assert at.text_input(key="company_name").value == "Second"
    assert "pma" not in at.session_state["_source_files"]
    # A restored source can be removed without leaving a hidden upload active.
    at.button(key="remove_curve").click().run()
    assert_ok(at)
    assert not at.session_state["_source_files"]


def test_pma_only_dossier_can_be_saved(monkeypatch):
    captured = []
    original = dossier_ui.encode_dossier
    def capture(*args, **kwargs):
        raw = original(*args, **kwargs); captured.append(raw); return raw
    monkeypatch.setattr(dossier_ui, "encode_dossier", capture)
    at = launch({"pma_subscribed_kva": 42.0}, {"pma": pma_file()})
    assert_ok(at)
    assert at.number_input(key="pma_subscribed_kva").value == 42.0
    assert decode_dossier(captured[-1])["files"] == {"pma": pma_file()}


def test_navigation_preserves_notes_and_settings(monkeypatch):
    monkeypatch.setattr("requests.get", Mock(side_effect=AssertionError("Offline")))
    at = launch({"company_name": "Navigation", "unique_electricity_price": 0.31})
    assert_ok(at)
    at.button(key="back_to_analysis").click().run()
    assert_ok(at)
    assert any("00h" in m.value and "23h" in m.value for m in at.markdown)
    at.text_area(key="note_profile").set_value("NOTE_NAVIGATION : vérifier les horaires.").run()
    at.button(key="prepare_report").click().run()
    assert_ok(at)
    assert at.text_area(key="note_profile").value.startswith("NOTE_NAVIGATION")
    text = "\n".join(p.extract_text() for p in PdfReader(BytesIO(at.session_state["_test_result"]["energy_pdf"])).pages)
    assert "NOTE_NAVIGATION" in text
    assert at.text_input(key="company_name").value == "Navigation"
    at.button(key="back_to_analysis").click().run()
    assert_ok(at)
    assert at.text_area(key="note_profile").value.startswith("NOTE_NAVIGATION")
    assert at.session_state["unique_electricity_price"] == 0.31


def test_solar_is_additive_and_saved_dossiers_restore_the_option(monkeypatch):
    monkeypatch.setattr("requests.get", Mock(return_value=solar_response()))
    at = launch({"analysis_mode": "Opportunité photovoltaïque", "manual_coordinates": True})
    assert_ok(at)
    assert at.checkbox(key="_include_pv").value is True
    at.button(key="back_to_analysis").click().run()
    assert_ok(at)
    core = {"Vue d’ensemble", "Profils", "Puissance", "Tarification", "Détail journalier", "Qualité des données"}
    assert core | {"Production solaire", "Simulation financière"} == {t.label for t in at.tabs}
    assert not any(r.label in {"Navigation", "Parcours"} for r in at.radio)
    at.checkbox(key="_include_pv").uncheck().run()
    assert_ok(at)
    assert {t.label for t in at.tabs} == core
    at.checkbox(key="_include_pv").check().run()
    assert_ok(at)
    assert core | {"Production solaire", "Simulation financière"} == {t.label for t in at.tabs}
