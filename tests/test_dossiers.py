import json
from datetime import date, time

import pytest

from dossiers import decode_dossier, encode_dossier, FORMAT


def test_roundtrip_keeps_files_settings_dates_times_and_notes():
    files = {"curve": {"name": "consommation.csv", "content": b"Horodate;Valeur\n"},
             "pma": {"name": "pma.xlsx", "content": bytes(range(256))}}
    fields = {"company_name": "Atelier électricité", "diagnostic_date": date(2026, 9, 10),
              "hc2_start": time(12, 30), "hc_range_count": 2, "pv_peak_kwp": 27.5,
              "annual_subscription_eur": 730.5, "note_context": "Froid & cuisson\nÀ confirmer",
              "report_status": "À compléter après rendez-vous"}
    restored = decode_dossier(encode_dossier(fields, files))
    assert restored["files"] == files
    assert all(restored["fields"][key] == value for key, value in fields.items())
    assert restored["fields"]["orientation_label"] == "Sud"


def test_legacy_v1_remains_readable():
    restored = decode_dossier(json.dumps({"format": FORMAT, "version": 1,
        "fields": {"note_context": "Ancienne note", "diagnostic_date": "2026-09-09"}}).encode())
    assert restored["version"] == 1
    assert restored["fields"]["diagnostic_date"] == date(2026, 9, 9)


@pytest.mark.parametrize("change", [
    lambda p: p.update(version=99),
    lambda p: p["fields"].update(load_curve_file="injection"),
    lambda p: p["fields"].update(pv_peak_kwp=-10),
    lambda p: p["fields"].update(hc1_start="29:10"),
    lambda p: p["fields"].update(annual_subscription_eur=float("nan")),
    lambda p: p["fields"].update(start_date="2026-02-03", end_date="2026-01-01"),
    lambda p: p["files"]["curve"].update(data="broken"),
    lambda p: p["files"]["curve"].update(sha256="incorrect"),
    lambda p: p["files"]["curve"].update(name="../outside.csv"),
])
def test_invalid_dossier_is_rejected(change):
    payload = json.loads(encode_dossier({}, {"curve": {"name": "test.csv", "content": b"a"}}))
    change(payload)
    with pytest.raises(ValueError): decode_dossier(json.dumps(payload).encode())
