"""Propositions issues du texte déjà lu, sans OCR supplémentaire ni cache."""
import re

FIELDS = {
    'company_name': "Nom du client / entreprise",
    'invoice_address': "Adresse affichée sur la facture (nature à vérifier)",
    'invoice_prm': "PDL / PRM (14 chiffres)",
    'invoice_date': "Date de facture",
    'invoice_period_start': "Début de période facturée",
    'invoice_period_end': "Fin de période facturée",
    'invoice_consumption': "Consommation totale (valeur et unité)",
    'invoice_power': "Puissance contractualisée (valeur et unité)",
    'invoice_total_ht': "Montant total hors TVA (€)",
    'invoice_total_ttc': "Montant total TTC (€)",
}


def extract_metadata(text):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    found = {key: [] for key in FIELDS}
    def add(key, value, source):
        found[key].append((value.strip(), source))
    date = r'\d{2}/\d{2}/\d{4}'
    amount = r'\d+(?:[ \u00a0\u202f]\d{3})*(?:[,.]\d+)'
    # Only standalone postal blocks; do not harvest supplier legal footers.
    first_page = text.split('--- Page 2 ---')[0]
    address_lines = [line.strip() for line in first_page.splitlines() if line.strip()]
    for i, line in enumerate(address_lines):
        if i and re.fullmatch(r'\d{5}\s+[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ \-\']+', line):
            street = address_lines[i-1]
            if re.match(r'\d+\s+(?:bis\s+|ter\s+)?(?:rue|avenue|boulevard|route|chemin|impasse|place|allée|allee|cours|quai)\b', street, re.I):
                add('invoice_address', street + ', ' + line, street + '\n' + line)
                if i >= 2:
                    name = address_lines[i-2]
                    if re.fullmatch(r'[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ &\-\'.]{1,79}', name) and not re.search(r'facture|adresse|livraison|facturation|client|page', name, re.I):
                        add('company_name', name, name + '\n' + street + '\n' + line)
    for i, line in enumerate(lines):
        context = line + (' ' + lines[i+1] if i+1 < len(lines) else '')
        if re.search(r'FACTURE.*(?:DU|DATE)', line, re.I):
            m = re.search(date, context)
            if m: add('invoice_date', m[0], context)
        m = re.search(r'\b(?:PDL|PRM|point de livraison)\s*[:n°º .-]*\s*(\d{14})(?!\d)', line, re.I)
        if m: add('invoice_prm', m[1], line)
        m = re.search(rf'\bdu\s+({date})\s+au\s+({date})', line, re.I)
        if m and re.search(r'période|periode|HPE|HCE|HPH|HCH|consommation', line, re.I):
            add('invoice_period_start', m[1], line)
            add('invoice_period_end', m[2], line)
        for key, label in [('invoice_total_ht', r'(?:MONTANT\s+)?TOTAL\s+(?:HT\b|HORS\s+TVA)'), ('invoice_total_ttc', r'(?:MONTANT\s+)?TOTAL\s+TTC')]:
            if re.search(label, line, re.I):
                m = re.search(rf'({amount})\s*€', line)
                if m: add(key, m[1], line)
        if re.search(r'consommation\s+(?:totale|réelle|reelle)', line, re.I):
            m = re.search(rf'({amount})\s*(MWh|kWh)\b', context, re.I)
            if m: add('invoice_consumption', m[1] + ' ' + m[2], context)
        # Do not infer kVA from an ambiguous kW/kVA header or an unlabelled table.
        m = re.search(rf'puissance\s+(?:souscrite|contractualisée|contractuelle)\s*:?\s*({amount}|\d+)\s*(kVA|kW)\b', line, re.I)
        if m: add('invoice_power', m[1] + ' ' + m[2], line)
    result = {}
    for key, entries in found.items():
        values = {value for value, _ in entries}
        if len(values) == 1:
            result[key] = {'value': entries[0][0], 'evidence': entries[0][1]}
    return result
