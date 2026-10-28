"""Successive meter files must not accumulate indefinitely in the global cache."""
import ast
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]


def test_four_imports_evict_old_file_but_reuse_current_file():
    names = {'normalize_enedis_columns', 'parse_enedis_dates',
             'read_enedis_file', 'read_geredis_file'}
    tree = ast.parse((ROOT / 'app.py').read_text())
    functions = '\n\n'.join(ast.unparse(n) for n in tree.body
                            if isinstance(n, ast.FunctionDef) and n.name in names)
    code = 'from io import BytesIO\nimport pandas as pd\nimport streamlit as st\n' + functions + '''
read_enedis_file.clear()
files = [("Horodate;Unité;Valeur;Pas\\n2024-09-03 00:05:00;W;" + str(i * 1000) + ";PT5M\\n").encode() for i in range(1, 5)]
frames = [read_enedis_file(data, "meter.csv") for data in files]
assert read_enedis_file(files[-1], "meter.csv").equals(frames[-1])
assert read_enedis_file(files[0], "meter.csv").equals(frames[0])
assert [float(frame.Valeur.iloc[0]) for frame in frames] == [1000, 2000, 3000, 4000]
read_enedis_file.clear()
'''
    with patch.object(pd, 'read_csv', wraps=pd.read_csv) as reader:
        at = AppTest.from_string(code, default_timeout=30).run()
    assert not at.exception, [e.message for e in at.exception]
    # Four initial parses + the evicted first file. Reusing the last is a hit.
    assert reader.call_count == 5


def test_all_data_caches_have_size_and_expiry_limits():
    tree = ast.parse((ROOT / 'app.py').read_text())
    checked = []
    for node in tree.body:
        if not isinstance(node, ast.FunctionDef):
            continue
        for decorator in node.decorator_list:
            if isinstance(decorator, ast.Call) and ast.unparse(decorator.func) == 'st.cache_data':
                options = {kw.arg: ast.literal_eval(kw.value) for kw in decorator.keywords}
                assert options['max_entries'] > 0
                assert options['ttl'] > 0
                checked.append(node.name)
    assert len(checked) == 5
