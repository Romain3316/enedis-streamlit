"""Palette historique commune à l'interface et aux exports CMA."""

import math
import numpy as np
from reportlab.lib import colors

COLOR_SCALE = [(0.00, "#63BE7B"), (0.30, "#A9D26D"), (0.50, "#FFEB84"),
               (0.72, "#F6B26B"), (1.00, "#F8696B")]
MISSING_COLOR = "#F2F5F7"


def bounds(matrix):
    values = matrix.select_dtypes(include=[np.number]).to_numpy(dtype=float)
    finite = values[np.isfinite(values)]
    return (float(finite.min()), float(finite.max())) if finite.size else (0.0, 1.0)


def cell_color(value, low, high):
    if value is None or not math.isfinite(float(value)):
        return MISSING_COLOR
    ratio = 0.5 if high <= low else min(1.0, max(0.0, (float(value)-low)/(high-low)))
    for (p0, h0), (p1, h1) in zip(COLOR_SCALE, COLOR_SCALE[1:]):
        if ratio <= p1:
            t = (ratio-p0)/(p1-p0)
            rgb0 = tuple(int(h0[i:i+2], 16) for i in (1,3,5))
            rgb1 = tuple(int(h1[i:i+2], 16) for i in (1,3,5))
            return '#' + ''.join(f'{round(a+t*(b-a)):02X}' for a,b in zip(rgb0,rgb1))
    return COLOR_SCALE[-1][1]


def pdf_cell_styles(matrix):
    """Use original, unrounded values and one common range for the whole matrix."""
    low, high = bounds(matrix)
    return [("BACKGROUND", (c+1,r+1), (c+1,r+1), colors.HexColor(cell_color(value,low,high)))
            for r, row in enumerate(matrix.to_numpy(dtype=float)) for c,value in enumerate(row)]


def style_matrix(matrix):
    low, high = bounds(matrix)
    numeric = list(matrix.select_dtypes(include=[np.number]).columns)
    styled = matrix.style.format({c: lambda v: '' if not math.isfinite(float(v)) else f'{v:.1f}'.replace('.',',') for c in numeric})
    if numeric:
        styled = styled.map(lambda v: f'background-color: {cell_color(v,low,high)}; color: #17365D', subset=numeric)
    return styled.set_properties(**{'text-align':'center', 'border':'1px solid white'})
