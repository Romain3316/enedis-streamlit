import pandas as pd
from reportlab.lib import colors
from heatmap_colors import COLOR_SCALE, MISSING_COLOR, cell_color, pdf_cell_styles, style_matrix


def test_report_and_interface_use_same_scale_on_unrounded_values():
    matrix = pd.DataFrame({'Lundi': [0.001, 0.003, float('nan')], 'Mardi': [0.002, 0.004, 0.005]})
    html = style_matrix(matrix).to_html()
    styles = pdf_cell_styles(matrix)
    for r, row in enumerate(matrix.to_numpy()):
        for c, value in enumerate(row):
            expected = cell_color(value, 0.001, 0.005)
            assert expected in html
            assert styles[r*2+c] == ('BACKGROUND', (c+1,r+1), (c+1,r+1), colors.HexColor(expected))
    assert cell_color(0.001, .001, .005) == COLOR_SCALE[0][1]
    assert cell_color(.005, .001, .005) == COLOR_SCALE[-1][1]
    assert cell_color(float('nan'), 0, 1) == MISSING_COLOR
    assert cell_color(5,5,5) == '#FFEB84'
