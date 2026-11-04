from io import BytesIO
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from invoice_reader import parse_invoice_text, read_invoice


TEXT = '''Coût de la Fourniture
HPE du 23/08/2025 au 21/09/2025 6,101 MWh 97,66 €/MWh 595,82 € 20%
HCE du 23/08/2025 au 21/09/2025 1714 kWh 0,08285 €/kWh 142,00 € 20%
Composante de soutirage variable HPE 6,101 MWh 21,30 €/MWh 129,95 €
'''


def test_invoice_units_and_ambiguous_prices():
    parsed = parse_invoice_text(TEXT)
    assert parsed['candidates']['HPE']['price'] == pytest.approx(0.09766)
    assert parsed['candidates']['HCE']['price'] == pytest.approx(0.08285)
    assert sum(r['kwh'] for r in parsed['rows']) == 7815
    assert set(parsed['candidates']) == {'HPE', 'HCE'}
    assert not parse_invoice_text(TEXT + TEXT)['candidates']
    assert not parse_invoice_text('Puissance souscrite 42 kW')['candidates']
    with pytest.raises(ValueError):
        read_invoice(b'not a pdf')
    with pytest.raises(ValueError):
        read_invoice(b'%PDF-' + b'0' * (10 * 1024 * 1024))


def test_digital_pdf_reading():
    from reportlab.pdfgen.canvas import Canvas
    buffer = BytesIO()
    canvas = Canvas(buffer)
    canvas.drawString(40, 700, TEXT.splitlines()[1])
    canvas.drawString(40, 670, 'Synthetic invoice for extraction testing without customer information.')
    canvas.save()
    result = read_invoice(buffer.getvalue())
    assert result['ocr_pages'] == []
    assert result['candidates']['HPE']['price'] == pytest.approx(0.09766)


def test_invoice_workspace_does_not_change_settings_without_confirmation(monkeypatch):
    import invoice_ui
    monkeypatch.setattr(invoice_ui.st, 'file_uploader', lambda *a, **kw: BytesIO(b'%PDF-test'))
    calls = []
    def read(data):
        calls.append(data)
        return dict(text=TEXT, pages=1, ocr_pages=[], **parse_invoice_text(TEXT))
    monkeypatch.setattr(invoice_ui, 'read_invoice', read)
    app = Path(__file__).resolve().parents[1] / 'app.py'
    at = AppTest.from_file(str(app), default_timeout=30)
    at.session_state['_workspace_page'] = 'Factures — Tests'
    at.session_state['hp_summer_electricity_price'] = 0.18
    at.session_state['pma_subscribed_kva'] = 36.0
    at.run()
    assert not at.exception
    assert not calls
    at.button(key='read_invoice').click().run()
    assert not at.exception
    assert len(calls) == 1
    assert at.session_state['hp_summer_electricity_price'] == 0.18
    at.multiselect(key='_invoice_selected').set_value(['HPE'])
    submit = next(b for b in at.button if b.label.startswith('Appliquer les valeurs'))
    submit.click().run()
    assert at.error
    assert at.session_state['hp_summer_electricity_price'] == 0.18
    at.checkbox(key='_invoice_verified').check()
    next(b for b in at.button if b.label.startswith('Appliquer les valeurs')).click().run()
    assert not at.exception and not at.error
    assert at.session_state['hp_summer_electricity_price'] == pytest.approx(0.09766)
    assert at.session_state['pma_subscribed_kva'] == 36.0
    assert at.session_state['electricity_tariff_type'] == 'HP / HC hiver-été'
    assert len(calls) == 1
    # Replacing a dossier must erase extracted invoice data, not carry it to another customer.
    from dossiers import encode_dossier, decode_dossier
    at.session_state['_pending_dossier'] = decode_dossier(encode_dossier({}, {}))
    at.run()
    assert not at.exception
    assert '_invoice_result' not in at.session_state
