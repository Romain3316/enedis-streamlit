import base64
from ui_theme import apply_theme, compact_header, dossier_summary, heat_legend, open_report, open_analysis, update_solar_option, render_profile_matrix
from heatmap_colors import style_matrix, pdf_cell_styles, cell_color, bounds, COLOR_SCALE
import json
from dossier_schema import SETTINGS, PV_KEYS
from dossier_ui import initialize_dossier, render_dossier_loader, source_uploader, render_dossier_download
from io import BytesIO
from pathlib import Path

import requests
from PIL import Image as PILImage, ImageDraw, ImageFont
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    KeepTogether,
    PageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from pvlib.location import Location


# ============================================================
# CONFIGURATION GÉNÉRALE
# ============================================================

st.set_page_config(
    page_title="CMA - Analyse énergétique",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)

CMA_BLUE = "#17365D"
CMA_RED = "#E53935"
CMA_GREY = "#F3F5F7"
CMA_TEXT = "#202735"

WEEKDAYS = {
    0: "Lundi",
    1: "Mardi",
    2: "Mercredi",
    3: "Jeudi",
    4: "Vendredi",
    5: "Samedi",
    6: "Dimanche",
}
WEEKDAY_ORDER = list(WEEKDAYS.values())

MONTHS = {
    1: "Janvier",
    2: "Février",
    3: "Mars",
    4: "Avril",
    5: "Mai",
    6: "Juin",
    7: "Juillet",
    8: "Août",
    9: "Septembre",
    10: "Octobre",
    11: "Novembre",
    12: "Décembre",
}


# ============================================================
# STYLE CMA
# ============================================================

st.markdown(
    """
    <style>
        :root {
            --cma-blue: #17365D;
            --cma-blue-dark: #102947;
            --cma-red: #E53935;
            --cma-red-dark: #C82E2A;
