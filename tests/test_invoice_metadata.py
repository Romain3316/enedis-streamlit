from invoice_metadata import extract_metadata, FIELDS
from streamlit.testing.v1 import AppTest

TEXT = '''FACTURE ELECTRICITE DU N° DE FACTURE
25/09/2025 001000991122
ATELIER TEST
12 rue des Fleurs
33000 BORDEAUX
HPE du 23/08/2025 au 21/09/2025 6,101 MWh 97,66 €/MWh
MONTANT TOTAL HORS TVA 20,0 % 1244,25 €
MONTANT TOTAL TTC A PAYER 1493,10 €
PRM : 12345678901234
Puissance souscrite : 42 kW
Consommation totale : 7,815 MWh
'''


def test_metadata_labels_ambiguity_and_no_invented_units():
    r = extract_metadata(TEXT)
    assert r['company_name']['value'] == 'ATELIER TEST'
    assert r['invoice_address']['value'] == '12 rue des Fleurs, 33000 BORDEAUX'
    assert r['invoice_total_ht']['value'] == '1244,25'
    assert r['invoice_date']['value'] == '25/09/2025'
    assert r['invoice_period_end']['value'] == '21/09/2025'
    assert r['invoice_power']['value'] == '42 kW'
    assert 'invoice_prm' not in extract_metadata('IBAN FR7612345678901234')
    assert 'invoice_power' not in extract_metadata('Puissance souscrite actuelle (kW ou kVA) : 36,0')
    assert 'invoice_address' not in extract_metadata(TEXT + '\nAUTRE CLIENT\n15 rue du Port\n33100 BORDEAUX')
    assert 'invoice_total_ht' not in extract_metadata(TEXT + '\nTOTAL HORS TVA 1400,00 €')


def test_validation_selective_apply_address_and_roundtrip():
    from dossiers import encode_dossier, decode_dossier
    at = AppTest.from_string('from invoice_ui import render_invoice_information\nrender_invoice_information(' + repr(TEXT) + ')')
    at.session_state['company_name'] = 'Ancien nom'
    at.session_state['company_address'] = 'Ancien site'
    at.session_state['pma_subscribed_kva'] = 36.0
    at.run()
    assert not at.exception
    assert at.multiselect(key='_invoice_info_selected').value == [FIELDS['company_name'], FIELDS['invoice_address']]
    assert at.text_input(key='_invoice_info_company_name').value == 'ATELIER TEST'
    assert at.session_state['company_name'] == 'Ancien nom'
    at.multiselect(key='_invoice_info_selected').set_value([FIELDS[k] for k in ('company_name', 'invoice_address', 'invoice_power')])
    at.button[0].click().run()
    assert at.error and at.session_state['company_name'] == 'Ancien nom'
    at.checkbox(key='_invoice_info_verified').check()
    at.button[0].click().run()
    assert not at.exception and not at.error
    assert at.session_state['company_name'] == 'ATELIER TEST'
    assert at.session_state['company_address'] == 'Ancien site'
    assert at.session_state['pma_subscribed_kva'] == 36.0
    at.session_state['address_candidates'] = [{'label': 'old'}]
    at.session_state['manual_coordinates'] = True
    at.checkbox(key='_invoice_info_site').check()
    at.button[0].click().run()
    assert not at.exception
    assert at.session_state['company_address'] == '12 rue des Fleurs, 33000 BORDEAUX'
    assert 'address_candidates' not in at.session_state
    assert at.session_state['manual_coordinates'] is False
    fields = {k: at.session_state[k] for k in ('company_name', 'company_address', 'invoice_address', 'invoice_power')}
    restored = decode_dossier(encode_dossier(fields, {}))
    assert all(restored['fields'][k] == v for k, v in fields.items())
