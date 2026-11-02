"""Horaires réels, unités et absence de rétention entre graphiques PMA."""
import ast
import gc
from pathlib import Path
import weakref

import pandas as pd
import plotly.graph_objects as go

from test_enedis_import import FUNCTIONS


def test_pma_hour_scatter():
    source = Path(__file__).resolve().parents[1] / "app.py"
    node = next(n for n in ast.parse(source.read_text()).body
                if isinstance(n, ast.FunctionDef) and n.name == "build_pma_hour_scatter")
    namespace = {"go": go, "pd": pd}
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(source), "exec"), namespace)
    build = namespace[node.name]
    raw = ("Horodate;Grandeur physique;Valeur;Unité\n"
           "2026-06-03 06:13:11;PMA;24950;VA\n"
           "2026-06-02 23:59:59;PMA;28;kVA\n"
           "2026-06-03 00:00:00;PMA1;5000;VA\n"
           "2026-06-03 08:01:02;PMA2;4000;VA\n"
           "2026-06-03 09:02:03;PMA3;3000;VA\n").encode()
    frame = FUNCTIONS["read_enedis_pma_file"](raw, "synthetic.csv")
    before = frame.copy(deep=True)
    fig = build(frame)
    assert [t.name for t in fig.data] == ["PMA", "PMA1", "PMA2", "PMA3"]
    assert list(fig.data[0].x) == [86399, 22391]
    assert list(fig.data[0].y) == [28, 24.95]
    assert fig.data[0].text[1] == "03/06/2026 à 06:13:11"
    assert fig.layout.xaxis.range == (0, 86400)
    assert fig.layout.xaxis.ticktext[-1] == "24:00"
    pd.testing.assert_frame_equal(frame, before)
    assert len(build(frame.iloc[:1]).data) == 1
    # Des reruns successifs ne doivent garder ni les figures ni leurs tableaux source.
    refs = []
    for _ in range(10):
        data = frame.copy()
        chart = build(data)
        refs.extend([weakref.ref(data), weakref.ref(chart)])
        del data, chart
    gc.collect()
    assert all(ref() is None for ref in refs)
