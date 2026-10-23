"""Régressions de lecture, avec données fictives uniquement.

Exécution autonome : python tests/test_enedis_import.py
"""
import ast
from io import BytesIO
from pathlib import Path
import unittest
from unittest.mock import patch

import pandas as pd


# Charger les fonctions réelles sans lancer l'interface Streamlit.
SOURCE = Path(__file__).resolve().parents[1] / "app.py"
NAMES = {
    "normalize_enedis_columns", "parse_enedis_dates", "read_enedis_file",
    "read_enedis_pma_file", "read_geredis_file", "detect_time_step",
    "detect_source_unit", "parse_pas_hours", "enrich_energy_data",
}
FUNCTIONS = {"pd": pd, "BytesIO": BytesIO}
for node in ast.parse(SOURCE.read_text(encoding="utf-8")).body:
    if isinstance(node, ast.FunctionDef) and node.name in NAMES:
        node.decorator_list = []
        exec(compile(ast.Module(body=[node], type_ignores=[]), str(SOURCE), "exec"), FUNCTIONS)


def analyze(raw, filename="synthetic.csv"):
    df = FUNCTIONS["read_enedis_file"](raw, filename)
    return FUNCTIONS["enrich_energy_data"](
        df, FUNCTIONS["detect_time_step"](df), FUNCTIONS["detect_source_unit"](df)
    )[0]


class EnedisImportTests(unittest.TestCase):
    def test_headers_encodings_and_five_minute_power(self):
        for header, encoding in [("Unité", "utf-8-sig"), ("Unité", "latin-1"),
                                 ("UnitÃ©", "utf-8"), ("Unite", "utf-8")]:
            with self.subTest(header=header, encoding=encoding):
                raw = (f"Grandeur physique;{header};Horodate;Valeur;Pas\n"
                       "PRI;VAr;2024-09-03 00:05:00;9000;PT5M\n"
                       "PA;W;2024-09-03 00:05:00;6000;PT5M\n"
                       "PA;W;2024-09-03 00:10:00;12000;PT5M\n").encode(encoding)
                result = analyze(raw)
                self.assertEqual(len(result), 2)
                self.assertAlmostEqual(result.Energie_kWh.sum(), 1.5)
                self.assertEqual(result.Puissance_kW.tolist(), [6, 12])
                self.assertEqual(result.Horodate.iloc[0], pd.Timestamp("2024-09-03 00:05"))

    def test_dates_are_independent_of_first_row(self):
        values = pd.Series(["2024-09-03 00:05:00", "2024-09-25 00:05:00",
                            "2024-10-27 02:05:00", "2024-10-27 02:05:00"])
        expected = pd.to_datetime(values, format="%Y-%m-%d %H:%M:%S")
        for order in [values, values.iloc[::-1]]:
            actual = FUNCTIONS["parse_enedis_dates"](order)
            pd.testing.assert_series_equal(actual, expected.loc[order.index])

    def test_french_and_native_excel_dates(self):
        for values in [pd.Series(["03/09/2024 01:00", "03/09/2024 02:00:00"]),
                       pd.Series(pd.date_range("2024-09-03 01:00", periods=2, freq="h"))]:
            df = pd.DataFrame({"Horodate": values, "Unité": "kW", "Valeur": 6})
            with patch.object(pd, "read_excel", return_value=df):
                result = analyze(b"", "synthetic.xlsx")
            self.assertEqual(result.Horodate.iloc[0], pd.Timestamp("2024-09-03 01:00"))
            self.assertAlmostEqual(result.Energie_kWh.sum(), 12)

    def test_mixed_steps_and_energy_units(self):
        raw = ("Horodate;Unité;Valeur;Pas\n"
               "2024-09-16 01:00:00;W;6000;PT60M\n"
               "2024-09-16 01:30:00;W;6000;PT30M\n"
               "2024-09-16 01:35:00;W;6000;PT5M\n")
        self.assertAlmostEqual(analyze(raw.encode()).Energie_kWh.sum(), 9.5)
        for unit, value in [("Wh", "500"), ("kWh", "0.5"), ("kW", "6")]:
            small = (f"Horodate;Unité;Valeur;Pas\n2024-09-03 00:05:00;{unit};{value};PT5M\n"
                     f"2024-09-03 00:10:00;{unit};{value};PT5M\n")
            self.assertAlmostEqual(analyze(small.encode()).Energie_kWh.sum(), 1)

    def test_invalid_dates_and_units_are_rejected(self):
        for values in [["2024-09-31 01:00:00"], ["invalid"], [None]]:
            with self.assertRaises(ValueError):
                FUNCTIONS["parse_enedis_dates"](pd.Series(values))
        for df in [pd.DataFrame({"Valeur": [1]}), pd.DataFrame({"Unité": [None]}),
                   pd.DataFrame({"Unité": [""]}), pd.DataFrame({"Unité": ["VA"]}),
                   pd.DataFrame({"Unité": ["W", "kW"]})]:
            with self.assertRaises(ValueError):
                FUNCTIONS["detect_source_unit"](df)
        with self.assertRaises(ValueError):
            FUNCTIONS["enrich_energy_data"](pd.DataFrame({"Valeur": [1]}), pd.Timedelta(minutes=5), "unknown")

    def test_pma_headers_and_dates(self):
        raw = ("Horodate;Grandeur physique;Valeur;UnitÃ©\n"
               "2024-09-03 12:00:00;PMA;25000;VA\n").encode()
        result = FUNCTIONS["read_enedis_pma_file"](raw, "synthetic-pma.csv")
        self.assertEqual(result.Puissance_kVA.iloc[0], 25)
        self.assertEqual(result.Horodate.iloc[0], pd.Timestamp("2024-09-03 12:00"))

    def test_geredis_still_uses_ten_minute_kw(self):
        result = analyze(b"03/09/2024 01:00 6 6 6 6 6 6\n", "synthetic.txt")
        self.assertEqual(len(result), 6)
        self.assertAlmostEqual(result.Energie_kWh.sum(), 6)


if __name__ == "__main__":
    unittest.main()
