import base64
from io import BytesIO
from pathlib import Path

import requests

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
from pvlib.location import Location


# ============================================================
# CONFIGURATION GÉNÉRALE
# ============================================================

st.set_page_config(
    page_title="CMA - Pré-diagnostic photovoltaïque",
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
            --cma-grey: #F3F5F7;
            --cma-border: #E1E7ED;
            --cma-text: #202735;
        }

        .stApp {
            background:
                radial-gradient(
                    circle at top right,
                    rgba(229, 57, 53, 0.035),
                    transparent 28rem
                ),
                #FFFFFF;
            color: var(--cma-text);
        }

        .block-container {
            max-width: 1500px;
            padding-top: 1rem;
            padding-bottom: 3rem;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(
                180deg,
                #F8F9FB 0%,
                #EEF2F6 100%
            );
            border-right: 1px solid #DDE3E9;
        }

        [data-testid="stSidebar"] .block-container {
            padding-top: 1.2rem;
        }

        /* Bandeau principal */
        .cma-header {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 48px;
            min-height: 220px;
            padding: 40px 44px;
            overflow: hidden;
            color: #FFFFFF;
            border-radius: 22px;
            margin-bottom: 22px;
            background: linear-gradient(
                116deg,
                var(--cma-blue) 0%,
                var(--cma-blue) 73%,
                var(--cma-red) 73%,
                var(--cma-red) 100%
            );
            box-shadow:
                0 20px 45px rgba(23, 54, 93, 0.18),
                0 4px 12px rgba(23, 54, 93, 0.09);
        }

        .cma-header::before {
            content: "☀";
            position: absolute;
            left: 58%;
            top: -54px;
            font-size: 215px;
            line-height: 1;
            color: rgba(255, 255, 255, 0.045);
            transform: rotate(-10deg);
            pointer-events: none;
        }

        .cma-header::after {
            content: "";
            position: absolute;
            right: 24%;
            bottom: -78px;
            width: 280px;
            height: 150px;
            border: 3px solid rgba(255, 255, 255, 0.05);
            border-radius: 24px;
            transform: rotate(-9deg);
            pointer-events: none;
        }

        .cma-header-content {
            position: relative;
            z-index: 2;
            max-width: 760px;
        }

        .cma-eyebrow {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 13px;
            padding: 7px 12px;
            border: 1px solid rgba(255, 255, 255, 0.28);
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.10);
            color: #FFFFFF;
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }

        .cma-header h1 {
            margin: 0;
            max-width: 720px;
            color: #FFFFFF !important;
            font-size: clamp(2.35rem, 4vw, 3.65rem);
            font-weight: 900;
            line-height: 1.02;
            letter-spacing: -0.045em;
            text-shadow: 0 4px 16px rgba(0, 0, 0, 0.28);
        }

        .cma-title-line {
            width: 98px;
            height: 5px;
            margin: 20px 0 17px;
            background: #FFFFFF;
            border-radius: 999px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
        }

        .cma-header p {
            margin: 0;
            max-width: 680px;
            color: rgba(255, 255, 255, 0.96);
            font-size: 1.2rem;
            font-weight: 500;
            line-height: 1.5;
        }

        /* Logo agrandi */
        .cma-logo {
            position: relative;
            z-index: 3;
            display: flex;
            align-items: center;
            justify-content: center;
            flex: 0 0 345px;
            width: 345px;
            min-height: 154px;
            padding: 16px 20px;
            background: #FFFFFF;
            color: var(--cma-blue);
            border: 1px solid rgba(255, 255, 255, 0.65);
            border-radius: 20px;
            text-align: center;
            font-size: 1.35rem;
            font-weight: 900;
            box-shadow:
                0 15px 32px rgba(0, 0, 0, 0.16),
                inset 0 0 0 1px rgba(23, 54, 93, 0.04);
        }

        .cma-logo img {
            display: block;
            width: 100%;
            max-width: 315px;
            max-height: 126px;
            object-fit: contain;
        }

        /* Introduction */
        .intro-card {
            display: flex;
            align-items: flex-start;
            gap: 15px;
            padding: 21px 24px;
            margin-bottom: 20px;
            background: #FFFFFF;
            border: 1px solid var(--cma-border);
            border-left: 6px solid var(--cma-red);
            border-radius: 16px;
            box-shadow: 0 8px 24px rgba(23, 54, 93, 0.07);
            font-size: 1.03rem;
            line-height: 1.65;
        }

        .intro-icon {
            display: flex;
            align-items: center;
            justify-content: center;
            flex: 0 0 42px;
            width: 42px;
            height: 42px;
            border-radius: 12px;
            background: #EDF3F9;
            font-size: 1.25rem;
        }

        /* Cartes d'accueil */
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 18px;
            margin: 8px 0 24px;
        }

        .feature-card {
            position: relative;
            min-height: 155px;
            padding: 25px;
            overflow: hidden;
            background: #FFFFFF;
            border: 1px solid var(--cma-border);
            border-radius: 18px;
            box-shadow: 0 8px 22px rgba(23, 54, 93, 0.07);
            transition:
                transform 0.22s ease,
                box-shadow 0.22s ease,
                border-color 0.22s ease;
        }

        .feature-card:hover {
            transform: translateY(-4px);
            border-color: rgba(23, 54, 93, 0.25);
            box-shadow: 0 16px 32px rgba(23, 54, 93, 0.12);
        }

        .feature-card::after {
            content: "";
            position: absolute;
            right: -28px;
            bottom: -35px;
            width: 110px;
            height: 110px;
            border-radius: 50%;
            background: rgba(23, 54, 93, 0.035);
        }

        .feature-icon {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 52px;
            height: 52px;
            margin-bottom: 16px;
            border-radius: 15px;
            background: #EAF1F8;
            font-size: 1.6rem;
        }

        .feature-card h3 {
            margin: 0 0 7px;
            color: var(--cma-blue) !important;
            font-size: 1.2rem;
            font-weight: 850;
        }

        .feature-card p {
            margin: 0;
            color: #687487;
            line-height: 1.5;
        }

        /* Cartes KPI */
        div[data-testid="stMetric"] {
            min-height: 122px;
            padding: 18px 19px;
            background: #FFFFFF;
            border: 1px solid var(--cma-border);
            border-top: 4px solid var(--cma-blue);
            border-radius: 15px;
            box-shadow: 0 7px 20px rgba(23, 54, 93, 0.06);
        }

        div[data-testid="stMetricLabel"] {
            color: #697589;
            font-size: 0.82rem;
            font-weight: 750;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }

        div[data-testid="stMetricValue"] {
            color: var(--cma-blue);
            font-size: 1.85rem;
            font-weight: 850;
        }

        button[data-baseweb="tab"] {
            font-weight: 750;
        }

        .stButton > button {
            min-height: 44px;
            padding: 0.65rem 1.15rem;
            background: var(--cma-red);
            color: #FFFFFF;
            border: 0;
            border-radius: 11px;
            font-weight: 750;
            box-shadow: 0 5px 13px rgba(229, 57, 53, 0.18);
        }

        .stButton > button:hover {
            background: var(--cma-red-dark);
            color: #FFFFFF;
            transform: translateY(-1px);
        }

        .stDownloadButton > button {
            min-height: 46px;
            background: var(--cma-blue);
            color: #FFFFFF;
            border: 0;
            border-radius: 11px;
            font-weight: 750;
            box-shadow: 0 5px 13px rgba(23, 54, 93, 0.16);
        }

        .stDownloadButton > button:hover {
            background: var(--cma-blue-dark);
            color: #FFFFFF;
            transform: translateY(-1px);
        }

        h2,
        h3 {
            color: var(--cma-blue);
        }


        .score-card {
            display: grid;
            grid-template-columns: 180px 1fr;
            gap: 24px;
            align-items: center;
            padding: 24px 26px;
            margin: 16px 0 22px;
            background: linear-gradient(135deg, #FFFFFF 0%, #F4F7FA 100%);
            border: 1px solid #DDE4EB;
            border-left: 7px solid var(--score-color, #17365D);
            border-radius: 18px;
            box-shadow: 0 10px 26px rgba(23, 54, 93, 0.09);
        }

        .score-circle {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 150px;
            height: 150px;
            margin: auto;
            border: 12px solid var(--score-color, #17365D);
            border-radius: 50%;
            background: #FFFFFF;
            box-shadow: inset 0 0 0 5px #F1F4F7;
        }

        .score-number {
            color: var(--score-color, #17365D);
            font-size: 2.55rem;
            line-height: 1;
            font-weight: 900;
        }

        .score-total {
            margin-top: 4px;
            color: #6B7688;
            font-size: 0.88rem;
            font-weight: 700;
        }

        .score-title {
            margin: 0 0 8px;
            color: var(--score-color, #17365D);
            font-size: 1.45rem;
            font-weight: 900;
        }

        .score-text {
            margin: 0;
            color: #3F4A5A;
            line-height: 1.58;
        }

        .pedagogy-card {
            padding: 17px 18px;
            margin: 10px 0 18px;
            background: #F7F9FB;
            border: 1px solid #E0E6EC;
            border-radius: 14px;
        }

        .pedagogy-card strong {
            color: #17365D;
        }

        .status-pill {
            display: inline-block;
            margin-top: 8px;
            padding: 5px 10px;
            border-radius: 999px;
            color: white;
            font-size: 0.78rem;
            font-weight: 800;
        }

        .score-breakdown {
            padding: 17px 19px;
            margin: 12px 0 20px;
            background: #FFFFFF;
            border: 1px solid #E1E7ED;
            border-radius: 14px;
        }

        @media (max-width: 1050px) {
            .cma-header {
                gap: 28px;
            }

            .cma-logo {
                flex-basis: 280px;
                width: 280px;
                min-height: 135px;
            }

            .cma-logo img {
                max-width: 250px;
                max-height: 108px;
            }
        }

        @media (max-width: 850px) {
            .cma-header {
                flex-direction: column;
                align-items: stretch;
                min-height: auto;
                padding: 30px 26px;
                background: linear-gradient(
                    165deg,
                    var(--cma-blue) 0%,
                    var(--cma-blue) 70%,
                    var(--cma-red) 70%,
                    var(--cma-red) 100%
                );
            }

            .cma-logo {
                width: 100%;
                max-width: 370px;
                flex-basis: auto;
                align-self: center;
            }

            .feature-grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# CORRECTIF V13 — THÈME CLAIR CMA VERROUILLÉ
# ============================================================

st.markdown(
    """
    <style>
        :root {
            color-scheme: light !important;
            --cma-blue: #17365D;
            --cma-blue-dark: #102947;
            --cma-red: #E53935;
            --cma-red-dark: #C82E2A;
            --cma-bg: #F5F7FA;
            --cma-surface: #FFFFFF;
            --cma-surface-soft: #F8FAFC;
            --cma-border: #DDE4EB;
            --cma-text: #202735;
            --cma-muted: #667287;
        }

        html,
        body,
        .stApp,
        [data-testid="stApp"],
        [data-testid="stAppViewContainer"] {
            color-scheme: light !important;
            background: var(--cma-bg) !important;
            color: var(--cma-text) !important;
        }

        [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(
                    circle at top right,
                    rgba(229,57,53,.035),
                    transparent 32rem
                ),
                linear-gradient(
                    180deg,
                    #FFFFFF 0%,
                    #F5F7FA 52rem
                ) !important;
        }

        [data-testid="stHeader"] {
            background: rgba(255,255,255,.96) !important;
            border-bottom: 1px solid var(--cma-border) !important;
        }

        h1, h2, h3, h4, h5, h6,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3 {
            color: var(--cma-blue) !important;
        }

        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li {
            color: var(--cma-text) !important;
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background:
                linear-gradient(
                    180deg,
                    #FFFFFF 0%,
                    #F2F5F8 100%
                ) !important;
            border-right: 1px solid var(--cma-border) !important;
        }

        [data-testid="stSidebar"] > div {
            background: transparent !important;
        }

        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span {
            color: var(--cma-blue) !important;
        }

        /* Conserver le bandeau sombre avec ses textes blancs */
        .cma-header,
        .cma-header * {
            color: #FFFFFF !important;
        }

        .cma-logo,
        .cma-logo * {
            color: var(--cma-blue) !important;
        }

        /* Cartes personnalisées */
        .intro-card,
        .feature-card,
        .score-card,
        .pedagogy-card,
        .score-breakdown {
            background-color: var(--cma-surface) !important;
            color: var(--cma-text) !important;
            border-color: var(--cma-border) !important;
        }

        .intro-card *,
        .feature-card *,
        .pedagogy-card *,
        .score-breakdown * {
            color: var(--cma-text) !important;
        }

        .feature-card h3,
        .pedagogy-card strong {
            color: var(--cma-blue) !important;
        }

        .feature-card p {
            color: var(--cma-muted) !important;
        }

        .score-title,
        .score-number {
            color: var(--score-color, var(--cma-blue)) !important;
        }

        .score-text,
        .score-total {
            color: var(--cma-muted) !important;
        }

        .status-pill {
            color: #FFFFFF !important;
        }

        /* KPI */
        [data-testid="stMetric"] {
            background: #FFFFFF !important;
            color: var(--cma-text) !important;
            border-color: var(--cma-border) !important;
        }

        [data-testid="stMetric"] * {
            color: var(--cma-text) !important;
        }

        [data-testid="stMetricLabel"] p,
        [data-testid="stMetricLabel"] div {
            color: var(--cma-muted) !important;
        }

        [data-testid="stMetricValue"] {
            color: var(--cma-blue) !important;
        }

        /* Onglets */
        [data-baseweb="tab-list"] {
            background: #FFFFFF !important;
            border: 1px solid var(--cma-border) !important;
            border-radius: 14px !important;
            padding: 5px !important;
        }

        button[data-baseweb="tab"] {
            color: var(--cma-muted) !important;
            background: transparent !important;
            border-radius: 10px !important;
        }

        button[data-baseweb="tab"] * {
            color: inherit !important;
        }

        button[data-baseweb="tab"]:hover {
            color: var(--cma-blue) !important;
            background: #EAF1F8 !important;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            color: var(--cma-red) !important;
            background: #FFF3F2 !important;
        }

        [data-baseweb="tab-highlight"] {
            background-color: var(--cma-red) !important;
        }

        /* Inputs, listes, dates, heures */
        input,
        textarea,
        [data-baseweb="input"] > div,
        [data-baseweb="base-input"],
        [data-baseweb="select"] > div {
            background: #FFFFFF !important;
            color: var(--cma-text) !important;
            -webkit-text-fill-color: var(--cma-text) !important;
            border-color: var(--cma-border) !important;
        }

        input::placeholder,
        textarea::placeholder {
            color: #98A2B1 !important;
            opacity: 1 !important;
        }

        [role="listbox"],
        [data-baseweb="popover"] > div,
        [data-baseweb="menu"] {
            background: #FFFFFF !important;
            color: var(--cma-text) !important;
        }

        [role="option"] {
            background: #FFFFFF !important;
            color: var(--cma-text) !important;
        }

        [role="option"]:hover,
        [role="option"][aria-selected="true"] {
            background: #EAF1F8 !important;
            color: var(--cma-blue) !important;
        }

        /* Radios, cases et curseurs */
        [role="radiogroup"] label,
        [role="radiogroup"] label *,
        [data-testid="stCheckbox"] label,
        [data-testid="stCheckbox"] label * {
            color: var(--cma-text) !important;
        }

        [data-testid="stSlider"] p,
        [data-testid="stSlider"] span {
            color: var(--cma-text) !important;
        }

        [data-testid="stSlider"] [role="slider"] {
            background: var(--cma-red) !important;
            border-color: var(--cma-red) !important;
        }

        /* Upload */
        [data-testid="stFileUploader"] section {
            background: #FFFFFF !important;
            color: var(--cma-text) !important;
            border: 1px dashed #AEB9C7 !important;
            border-radius: 14px !important;
        }

        [data-testid="stFileUploader"] section * {
            color: var(--cma-text) !important;
        }

        [data-testid="stFileUploader"] button {
            background: var(--cma-blue) !important;
            color: #FFFFFF !important;
            border: none !important;
        }

        [data-testid="stFileUploaderFile"] {
            color: var(--cma-text) !important;
            background: transparent !important;
        }

        /* Boutons */
        .stButton > button,
        .stDownloadButton > button {
            color: #FFFFFF !important;
        }

        .stButton > button * ,
        .stDownloadButton > button * {
            color: #FFFFFF !important;
        }

        /* Alertes */
        [data-testid="stAlert"] {
            color: var(--cma-text) !important;
            border: 1px solid var(--cma-border) !important;
            border-radius: 14px !important;
        }

        [data-testid="stAlert"] * {
            color: inherit !important;
        }

        /* Expanders */
        [data-testid="stExpander"] {
            background: #FFFFFF !important;
            border: 1px solid var(--cma-border) !important;
            border-radius: 14px !important;
        }

        [data-testid="stExpander"] summary,
        [data-testid="stExpander"] summary * {
            color: var(--cma-blue) !important;
            background: #FFFFFF !important;
        }

        /* Dataframes et tableaux */
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {
            background: #FFFFFF !important;
            border: 1px solid var(--cma-border) !important;
            border-radius: 14px !important;
            overflow: hidden !important;
        }

        [data-testid="stDataFrame"] iframe {
            color-scheme: light !important;
            background: #FFFFFF !important;
        }

        /* Plotly */
        [data-testid="stPlotlyChart"] {
            background: #FFFFFF !important;
            border: 1px solid var(--cma-border) !important;
            border-radius: 16px !important;
            padding: 7px !important;
            box-shadow: 0 7px 20px rgba(23,54,93,.05) !important;
            overflow: hidden !important;
        }

        [data-testid="stPlotlyChart"] > div {
            background: #FFFFFF !important;
            width: 100% !important;
            max-width: 100% !important;
        }

        [data-testid="stPlotlyChart"] .js-plotly-plot,
        [data-testid="stPlotlyChart"] .plot-container,
        [data-testid="stPlotlyChart"] .svg-container {
            width: 100% !important;
            max-width: 100% !important;
        }

        /* Chaque bloc Streamlit conserve sa hauteur réelle. */
        [data-testid="stVerticalBlock"] {
            min-width: 0 !important;
        }

        [data-testid="column"] {
            min-width: 0 !important;
            overflow: visible !important;
        }

        /* Infobulles */
        [role="tooltip"],
        [data-baseweb="tooltip"] {
            background: var(--cma-blue-dark) !important;
            color: #FFFFFF !important;
            border-radius: 9px !important;
        }

        [role="tooltip"] *,
        [data-baseweb="tooltip"] * {
            color: #FFFFFF !important;
        }

        @media (max-width: 850px) {
            [data-baseweb="tab-list"] {
                overflow-x: auto !important;
                justify-content: flex-start !important;
            }

            button[data-baseweb="tab"] {
                flex: 0 0 auto !important;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# CORRECTIF V13.1 — VARIABLES NATIVES STREAMLIT + MODE CLAIR
# ============================================================

st.markdown(
    """
    <style>
        /* Variables internes utilisées par Streamlit/BaseWeb */
        :root,
        html,
        body,
        .stApp,
        [data-testid="stApp"],
        [data-testid="stAppViewContainer"] {
            color-scheme: light !important;

            --primary-color: #E53935 !important;
            --background-color: #F5F7FA !important;
            --secondary-background-color: #FFFFFF !important;
            --text-color: #202735 !important;

            --bg-color: #F5F7FA !important;
            --fg-color: #202735 !important;
            --border-color: #DDE4EB !important;
            --base-radius: 0.75rem !important;
        }

        html,
        body {
            background: #F5F7FA !important;
        }

        .stApp,
        [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(
                    circle at top right,
                    rgba(229,57,53,.035),
                    transparent 32rem
                ),
                linear-gradient(
                    180deg,
                    #FFFFFF 0%,
                    #F5F7FA 52rem
                ) !important;
        }

        /* Évite les textes délavés injectés par le thème sombre */
        .stApp [data-testid="stMarkdownContainer"],
        .stApp [data-testid="stMarkdownContainer"] *,
        .stApp [data-testid="stWidgetLabel"],
        .stApp [data-testid="stWidgetLabel"] *,
        .stApp [data-testid="stCaptionContainer"],
        .stApp [data-testid="stCaptionContainer"] *,
        .stApp [data-testid="stMetric"],
        .stApp [data-testid="stMetric"] *,
        .stApp [data-baseweb="tab"],
        .stApp [data-baseweb="tab"] *,
        .stApp [role="radiogroup"],
        .stApp [role="radiogroup"] *,
        .stApp [data-testid="stFileUploader"],
        .stApp [data-testid="stFileUploader"] * {
            opacity: 1 !important;
        }

        /* Bandeaux info/success/warning/error */
        .stApp [data-testid="stAlert"] {
            opacity: 1 !important;
        }

        .stApp [data-testid="stAlert"] p,
        .stApp [data-testid="stAlert"] div,
        .stApp [data-testid="stAlert"] span {
            color: #17365D !important;
            -webkit-text-fill-color: #17365D !important;
            opacity: 1 !important;
        }

        .stApp [data-testid="stAlert"] svg {
            color: #17365D !important;
            fill: #17365D !important;
        }

        /* Captions */
        .stApp [data-testid="stCaptionContainer"],
        .stApp [data-testid="stCaptionContainer"] p,
        .stApp [data-testid="stCaptionContainer"] span {
            color: #667287 !important;
            -webkit-text-fill-color: #667287 !important;
            opacity: 1 !important;
        }

        /* KPI */
        .stApp [data-testid="stMetric"] {
            background: #FFFFFF !important;
        }

        .stApp [data-testid="stMetricLabel"],
        .stApp [data-testid="stMetricLabel"] p,
        .stApp [data-testid="stMetricLabel"] span {
            color: #667287 !important;
            -webkit-text-fill-color: #667287 !important;
            opacity: 1 !important;
        }

        .stApp [data-testid="stMetricValue"],
        .stApp [data-testid="stMetricValue"] div {
            color: #17365D !important;
            -webkit-text-fill-color: #17365D !important;
            opacity: 1 !important;
        }

        /* Tabs : neutraliser les opacités automatiques */
        .stApp [data-baseweb="tab-list"] {
            background: #FFFFFF !important;
        }

        .stApp button[data-baseweb="tab"] {
            opacity: 1 !important;
            color: #667287 !important;
            -webkit-text-fill-color: #667287 !important;
        }

        .stApp button[data-baseweb="tab"] * {
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
            opacity: 1 !important;
        }

        .stApp button[data-baseweb="tab"][aria-selected="true"] {
            color: #E53935 !important;
            -webkit-text-fill-color: #E53935 !important;
            background: #FFF3F2 !important;
        }

        /* Inputs et listes */
        .stApp input,
        .stApp textarea,
        .stApp [data-baseweb="input"] *,
        .stApp [data-baseweb="select"] * {
            opacity: 1 !important;
        }

        .stApp input,
        .stApp textarea {
            background: #FFFFFF !important;
            color: #202735 !important;
            -webkit-text-fill-color: #202735 !important;
        }

        /* Sidebar */
        .stApp [data-testid="stSidebar"] {
            background: linear-gradient(
                180deg,
                #FFFFFF 0%,
                #F2F5F8 100%
            ) !important;
        }

        .stApp [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        .stApp [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] *,
        .stApp [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
        .stApp [data-testid="stSidebar"] [data-testid="stWidgetLabel"] * {
            color: #17365D !important;
            -webkit-text-fill-color: #17365D !important;
            opacity: 1 !important;
        }

        /* Exception volontaire : textes blancs du bandeau */
        .stApp .cma-header,
        .stApp .cma-header *,
        .stApp .status-pill,
        .stApp .status-pill * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            opacity: 1 !important;
        }

        .stApp .cma-logo,
        .stApp .cma-logo * {
            color: #17365D !important;
            -webkit-text-fill-color: #17365D !important;
        }

        /* Fond des graphiques Plotly, même avant rendu JS */
        .stApp [data-testid="stPlotlyChart"],
        .stApp [data-testid="stPlotlyChart"] > div,
        .stApp [data-testid="stPlotlyChart"] .js-plotly-plot,
        .stApp [data-testid="stPlotlyChart"] .plot-container,
        .stApp [data-testid="stPlotlyChart"] .svg-container {
            background: #FFFFFF !important;
            color-scheme: light !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FONCTIONS
# ============================================================

_PLOTLY_CHART_ORIGINAL = st.plotly_chart


def _apply_cma_plotly_theme(figure):
    """Applique uniquement le thème graphique clair CMA."""
    try:
        figure.update_layout(
            template="plotly_white",
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(
                color="#202735",
                family="Arial, sans-serif",
            ),
            title_font=dict(
                color="#17365D",
                size=18,
            ),
            legend=dict(
                font=dict(color="#202735"),
                bgcolor="rgba(255,255,255,0.88)",
            ),
            hoverlabel=dict(
                bgcolor="#FFFFFF",
                bordercolor="#DDE4EB",
                font=dict(color="#202735"),
            ),
        )

        figure.update_xaxes(
            color="#202735",
            gridcolor="#E7ECF1",
            zerolinecolor="#DDE4EB",
            linecolor="#BFC9D4",
            tickfont=dict(color="#202735"),
            title_font=dict(color="#17365D"),
        )

        figure.update_yaxes(
            color="#202735",
            gridcolor="#E7ECF1",
            zerolinecolor="#DDE4EB",
            linecolor="#BFC9D4",
            tickfont=dict(color="#202735"),
            title_font=dict(color="#17365D"),
        )
    except Exception:
        pass

    return figure


def cma_plotly_chart(figure, *args, **kwargs):
    figure = _apply_cma_plotly_theme(figure)
    kwargs.setdefault(
        "config",
        {
            "displaylogo": False,
            "responsive": True,
        },
    )
    return _PLOTLY_CHART_ORIGINAL(
        figure,
        *args,
        **kwargs,
    )


st.plotly_chart = cma_plotly_chart



def file_to_base64(path: Path) -> str | None:
    if not path.exists():
        return None

    encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
    extension = path.suffix.lower().replace(".", "")
    mime = "jpeg" if extension in {"jpg", "jpeg"} else extension
    return f"data:image/{mime};base64,{encoded}"


def render_header() -> None:
    logo_paths = [
        Path("logo_cma.png"),
        Path("logo_cma.jpg"),
        Path("assets/logo_cma.png"),
        Path("assets/logo_cma.jpg"),
    ]

    logo_uri = None

    for logo_path in logo_paths:
        logo_uri = file_to_base64(logo_path)

        if logo_uri:
            break

    if logo_uri:
        logo_html = (
            '<div class="cma-logo">'
            f'<img src="{logo_uri}" alt="Logo CMA Nouvelle-Aquitaine">'
            "</div>"
        )
    else:
        logo_html = (
            '<div class="cma-logo">'
            'CMA<br><span style="font-size:0.76rem;font-weight:650;">'
            "NOUVELLE-AQUITAINE</span></div>"
        )

    st.markdown(
        f"""
        <div class="cma-header">
            <div class="cma-header-content">
                <div class="cma-eyebrow">
                    CMA Nouvelle-Aquitaine · Outil métier
                </div>
                <h1>Pré-diagnostic photovoltaïque</h1>
                <div class="cma-title-line"></div>
                <p>
                    Analyse automatisée des courbes de charge Enedis et
                    génération d'indicateurs d'aide au diagnostic énergétique.
                </p>
            </div>
            {logo_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def read_enedis_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    buffer = BytesIO(file_bytes)

    if filename.lower().endswith(".csv"):
        attempts = [
            {"sep": ";", "encoding": "utf-8"},
            {"sep": ";", "encoding": "latin-1"},
            {"sep": ",", "encoding": "utf-8"},
            {"sep": ",", "encoding": "latin-1"},
        ]

        df = None
        last_error = None

        for params in attempts:
            try:
                buffer.seek(0)
                candidate = pd.read_csv(buffer, **params)

                if {"Horodate", "Valeur"}.issubset(candidate.columns):
                    df = candidate
                    break
            except Exception as exc:
                last_error = exc

        if df is None:
            raise ValueError(
                "Le CSV ne contient pas les colonnes Horodate et Valeur "
                "ou son format n'est pas reconnu."
            ) from last_error
    else:
        df = pd.read_excel(buffer)

    required = {"Horodate", "Valeur"}
    missing = required - set(df.columns)

    if missing:
        raise ValueError(
            "Colonne(s) manquante(s) : " + ", ".join(sorted(missing))
        )

    available_columns = [
        column
        for column in ["Unité", "Horodate", "Valeur", "Nature", "Pas"]
        if column in df.columns
    ]

    df = df[available_columns].copy()

    df["Horodate"] = pd.to_datetime(
        df["Horodate"],
        errors="coerce",
        dayfirst=True,
    )

    df["Valeur"] = pd.to_numeric(
        df["Valeur"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )

    df = df.dropna(subset=["Horodate", "Valeur"])
    df = df.sort_values("Horodate").reset_index(drop=True)

    if df.empty:
        raise ValueError("Aucune donnée exploitable n'a été trouvée.")

    return df


def detect_time_step(df: pd.DataFrame) -> pd.Timedelta:
    differences = df["Horodate"].diff()
    positive_differences = differences[differences > pd.Timedelta(0)]

    if positive_differences.empty:
        raise ValueError("Impossible de détecter le pas de temps.")

    modes = positive_differences.mode()

    if not modes.empty:
        return modes.iloc[0]

    return positive_differences.median()


def detect_source_unit(df: pd.DataFrame) -> str:
    if "Unité" not in df.columns:
        return "Wh"

    units = (
        df["Unité"]
        .dropna()
        .astype(str)
        .str.strip()
    )
    units = units[units != ""]

    if units.empty:
        return "Wh"

    return units.mode().iloc[0]


def parse_pas_hours(value, default_step: pd.Timedelta) -> float:
    """Convertit PT30M, PT60M, PT1H... en durée exprimée en heures."""
    if pd.isna(value):
        return default_step.total_seconds() / 3600

    text = str(value).strip().upper()

    try:
        if text.startswith("PT") and text.endswith("M"):
            return float(text[2:-1]) / 60
        if text.startswith("PT") and text.endswith("H"):
            return float(text[2:-1])
    except ValueError:
        pass

    return default_step.total_seconds() / 3600


def enrich_energy_data(
    df: pd.DataFrame,
    time_step: pd.Timedelta,
    source_unit: str,
) -> tuple[pd.DataFrame, str]:
    result = df.copy()

    # Le fichier peut contenir plusieurs pas de temps (ex. PT60M puis PT30M).
    # On calcule donc la durée ligne par ligne quand la colonne Pas est présente.
    if "Pas" in result.columns:
        result["Duree_h"] = result["Pas"].apply(
            lambda value: parse_pas_hours(value, time_step)
        )
    else:
        result["Duree_h"] = time_step.total_seconds() / 3600

    normalized_unit = str(source_unit).lower().replace(" ", "")

    if normalized_unit in {"wh", "w.h"}:
        result["Energie_kWh"] = result["Valeur"] / 1000
        result["Puissance_kW"] = result["Energie_kWh"] / result["Duree_h"]
        message = "Les valeurs sont interprétées comme une énergie en Wh par intervalle."

    elif normalized_unit in {"kwh", "kw.h"}:
        result["Energie_kWh"] = result["Valeur"]
        result["Puissance_kW"] = result["Energie_kWh"] / result["Duree_h"]
        message = "Les valeurs sont interprétées comme une énergie en kWh par intervalle."

    elif normalized_unit in {"w", "watt", "watts"}:
        result["Puissance_kW"] = result["Valeur"] / 1000
        result["Energie_kWh"] = result["Puissance_kW"] * result["Duree_h"]
        message = "Les valeurs sont interprétées comme une puissance moyenne en W."

    elif normalized_unit in {"kw", "kilowatt", "kilowatts"}:
        result["Puissance_kW"] = result["Valeur"]
        result["Energie_kWh"] = result["Puissance_kW"] * result["Duree_h"]
        message = "Les valeurs sont interprétées comme une puissance moyenne en kW."

    else:
        result["Energie_kWh"] = result["Valeur"] / 1000
        result["Puissance_kW"] = result["Energie_kWh"] / result["Duree_h"]
        message = (
            f"Unité « {source_unit} » non reconnue : "
            "hypothèse Wh par intervalle."
        )

    return result, message


def filter_period(
    df: pd.DataFrame,
    period_mode: str,
    selected_year: int | None,
    start_date,
    end_date,
) -> pd.DataFrame:
    result = df.copy()

    if period_mode == "Année" and selected_year is not None:
        result = result[result["Horodate"].dt.year == selected_year]

    elif period_mode == "Période personnalisée":
        start_timestamp = pd.Timestamp(start_date)
        end_exclusive = pd.Timestamp(end_date) + pd.Timedelta(days=1)

        result = result[
            (result["Horodate"] >= start_timestamp)
            & (result["Horodate"] < end_exclusive)
        ]

    return result.copy()


def build_hourly_data(df: pd.DataFrame) -> pd.DataFrame:
    hourly = df.copy()

    # Convention Enedis : l'horodatage correspond à la FIN de l'intervalle.
    # Ainsi, pour un pas de 30 minutes :
    #   08:30 + 09:00 = heure se terminant à 09:00.
    # dt.ceil("h") conserve les heures rondes et rattache 08:30 à 09:00.
    hourly["Horodate_heure"] = hourly["Horodate"].dt.ceil("h")
    hourly["Puissance_ponderee"] = (
        hourly["Puissance_kW"] * hourly["Duree_h"]
    )

    hourly = (
        hourly.groupby("Horodate_heure", as_index=False)
        .agg(
            Energie_kWh=("Energie_kWh", "sum"),
            Puissance_ponderee=("Puissance_ponderee", "sum"),
            Duree_totale_h=("Duree_h", "sum"),
            Nombre_points=("Valeur", "size"),
        )
        .rename(columns={"Horodate_heure": "Horodate"})
        .sort_values("Horodate")
        .reset_index(drop=True)
    )

    hourly["Puissance_kW"] = (
        hourly["Puissance_ponderee"] / hourly["Duree_totale_h"]
    )

    return hourly[
        [
            "Horodate",
            "Energie_kWh",
            "Puissance_kW",
            "Nombre_points",
        ]
    ]


def build_daily_data(df: pd.DataFrame) -> pd.DataFrame:
    daily = df.copy()
    daily["Date"] = daily["Horodate"].dt.normalize()

    daily = (
        daily.groupby("Date", as_index=False)
        .agg(
            Consommation_kWh=("Energie_kWh", "sum"),
            Puissance_moyenne_kW=("Puissance_kW", "mean"),
            Puissance_max_kW=("Puissance_kW", "max"),
            Nombre_points=("Valeur", "size"),
        )
        .sort_values("Date")
        .reset_index(drop=True)
    )

    daily["Jour_num"] = daily["Date"].dt.weekday
    daily["Jour"] = daily["Jour_num"].map(WEEKDAYS)
    daily["Année"] = daily["Date"].dt.year
    daily["Mois_num"] = daily["Date"].dt.month
    daily["Mois"] = daily["Mois_num"].map(MONTHS)

    return daily


def build_monthly_data(df: pd.DataFrame) -> pd.DataFrame:
    monthly = df.copy()
    monthly["Mois_date"] = (
        monthly["Horodate"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    return (
        monthly.groupby("Mois_date", as_index=False)
        .agg(
            Consommation_kWh=("Energie_kWh", "sum"),
            Puissance_moyenne_kW=("Puissance_kW", "mean"),
            Puissance_max_kW=("Puissance_kW", "max"),
        )
        .sort_values("Mois_date")
        .reset_index(drop=True)
    )


def build_weekday_hour_matrix(hourly: pd.DataFrame) -> pd.DataFrame:
    data = hourly.copy()
    data["Jour_num"] = data["Horodate"].dt.weekday
    data["Jour"] = data["Jour_num"].map(WEEKDAYS)
    data["Heure"] = data["Horodate"].dt.hour

    matrix = data.pivot_table(
        index="Heure",
        columns="Jour",
        values="Puissance_kW",
        aggfunc="mean",
    )

    return matrix.reindex(
        index=range(24),
        columns=WEEKDAY_ORDER,
    )


def make_colored_style(matrix: pd.DataFrame):
    return (
        matrix.style
        .format("{:.1f}")
        .background_gradient(
            cmap="RdYlGn_r",
            axis=None,
        )
        .set_properties(
            **{
                "text-align": "center",
                "font-weight": "600",
                "border": "1px solid #FFFFFF",
            }
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", CMA_BLUE),
                        ("color", "white"),
                        ("text-align", "center"),
                        ("font-weight", "700"),
                    ],
                }
            ]
        )
    )


def build_daily_calendar(
    daily_df: pd.DataFrame,
    year: int,
) -> pd.DataFrame:
    year_data = daily_df[daily_df["Année"] == year].copy()

    if year_data.empty:
        return pd.DataFrame()

    first_day = pd.Timestamp(f"{year}-01-01")
    first_monday = first_day - pd.Timedelta(days=first_day.weekday())

    year_data["Index_jour"] = (
        year_data["Date"] - first_monday
    ).dt.days

    year_data["Semaine_calendrier"] = (
        year_data["Index_jour"] // 7 + 1
    )

    year_data["Jour_num"] = year_data["Date"].dt.weekday

    matrix = year_data.pivot_table(
        index="Semaine_calendrier",
        columns="Jour_num",
        values="Consommation_kWh",
        aggfunc="sum",
    )

    return matrix.reindex(columns=range(7))



@st.cache_data(ttl=86400, show_spinner=False)
def geocode_company_address(address: str) -> list[dict]:
    """
    Géocode une adresse avec le service officiel de la Géoplateforme.

    Paramètres attendus par l'API :
    - q : texte recherché ;
    - limit : nombre maximal de résultats ;
    - index : type de données recherché, ici les adresses.
    """
    address = address.strip()

    if len(address) < 5:
        raise ValueError("Veuillez saisir une adresse plus complète.")

    url = "https://data.geopf.fr/geocodage/search"
    params = {
        "q": address,
        "limit": 5,
        "index": "address",
    }
    headers = {
        "Accept": "application/json",
        "User-Agent": "CMA-Nouvelle-Aquitaine-SolarDiag/1.0",
    }

    response = requests.get(
        url,
        params=params,
        headers=headers,
        timeout=25,
    )

    if response.status_code == 400:
        raise ValueError(
            "La requête d'adresse a été refusée. Vérifiez que l'adresse "
            "contient au minimum un numéro, une voie et une commune."
        )

    response.raise_for_status()
    payload = response.json()

    candidates = []

    for feature in payload.get("features", []):
        properties = feature.get("properties", {})
        geometry = feature.get("geometry", {})
        coordinates = geometry.get("coordinates", [])

        if len(coordinates) < 2:
            continue

        longitude = float(coordinates[0])
        latitude = float(coordinates[1])

        label = (
            properties.get("label")
            or properties.get("name")
            or properties.get("display_name")
            or address
        )

        postcode = (
            properties.get("postcode")
            or properties.get("postalcode")
            or ""
        )
        city = (
            properties.get("city")
            or properties.get("municipality")
            or properties.get("locality")
            or ""
        )
        street = (
            properties.get("street")
            or properties.get("name")
            or ""
        )
        housenumber = properties.get("housenumber") or ""
        score = properties.get("score")

        candidates.append(
            {
                "label": str(label),
                "latitude": latitude,
                "longitude": longitude,
                "score": score,
                "postcode": str(postcode),
                "city": str(city),
                "street": str(street),
                "housenumber": str(housenumber),
                "source": "Géoplateforme / BAN",
            }
        )

    # Compatibilité avec certains formats de réponse alternatifs.
    for result in payload.get("results", []):
        longitude = (
            result.get("x")
            or result.get("lon")
            or result.get("longitude")
        )
        latitude = (
            result.get("y")
            or result.get("lat")
            or result.get("latitude")
        )

        if longitude is None or latitude is None:
            continue

        candidates.append(
            {
                "label": str(
                    result.get("fulltext")
                    or result.get("label")
                    or result.get("name")
                    or address
                ),
                "latitude": float(latitude),
                "longitude": float(longitude),
                "score": result.get("score"),
                "postcode": str(
                    result.get("postcode")
                    or result.get("postalcode")
                    or ""
                ),
                "city": str(
                    result.get("city")
                    or result.get("municipality")
                    or ""
                ),
                "street": str(result.get("street") or ""),
                "housenumber": str(
                    result.get("housenumber") or ""
                ),
                "source": "Géoplateforme / BAN",
            }
        )

    unique = []
    seen = set()

    for candidate in candidates:
        key = (
            round(candidate["latitude"], 6),
            round(candidate["longitude"], 6),
        )

        if key not in seen:
            seen.add(key)
            unique.append(candidate)

    if not unique:
        raise ValueError(
            "Aucune adresse précise n'a été trouvée. "
            "Ajoutez le numéro, la voie, le code postal et la commune."
        )

    return unique[:5]


def localize_paris(times: pd.Series | pd.DatetimeIndex) -> pd.DatetimeIndex:
    """Localise des horodatages Enedis naïfs dans le fuseau Europe/Paris."""
    index = pd.DatetimeIndex(pd.to_datetime(times))

    try:
        return index.tz_localize(
            "Europe/Paris",
            ambiguous="infer",
            nonexistent="shift_forward",
        )
    except Exception:
        return index.tz_localize(
            "Europe/Paris",
            ambiguous=True,
            nonexistent="shift_forward",
        )


def add_astronomical_solar_data(
    df: pd.DataFrame,
    latitude: float,
    longitude: float,
) -> pd.DataFrame:
    """
    Ajoute les informations solaires exactes pour chaque relevé.

    Les horodatages Enedis correspondent à la fin d'un intervalle.
    La position du soleil est donc calculée au milieu de chaque intervalle,
    ce qui évite de classer toute une heure comme nocturne alors que le soleil
    se lève ou se couche pendant cette heure.
    """
    result = df.copy()

    location = Location(
        latitude=latitude,
        longitude=longitude,
        tz="Europe/Paris",
    )

    if "Duree_h" not in result.columns:
        result["Duree_h"] = 1.0

    result["Horodate_milieu"] = (
        result["Horodate"]
        - pd.to_timedelta(result["Duree_h"] / 2, unit="h")
    )

    local_midpoints = localize_paris(result["Horodate_milieu"])
    solar_position = location.get_solarposition(local_midpoints)

    result["Hauteur_soleil_deg"] = pd.to_numeric(
        solar_position["apparent_elevation"].to_numpy(),
        errors="coerce",
    )

    # Le seuil astronomique classique du lever/coucher est proche de -0,833°,
    # afin de tenir compte du rayon apparent du soleil et de la réfraction.
    result["Soleil_leve"] = (
        result["Hauteur_soleil_deg"].fillna(-90) >= -0.833
    )

    # Calcul des événements solaires pour chaque date locale.
    result["Date_solaire"] = result["Horodate_milieu"].dt.normalize()
    unique_dates = pd.DatetimeIndex(
        result["Date_solaire"].dropna().drop_duplicates().sort_values()
    )

    local_noons = (
        unique_dates
        + pd.Timedelta(hours=12)
    ).tz_localize(
        "Europe/Paris",
        ambiguous=True,
        nonexistent="shift_forward",
    )

    sun_events = location.get_sun_rise_set_transit(
        local_noons,
        method="spa",
    )

    sun_table = pd.DataFrame(
        {
            "Date_solaire": unique_dates,
            "Lever_soleil": pd.DatetimeIndex(
                sun_events["sunrise"]
            )
            .tz_convert("Europe/Paris")
            .tz_localize(None),
            "Coucher_soleil": pd.DatetimeIndex(
                sun_events["sunset"]
            )
            .tz_convert("Europe/Paris")
            .tz_localize(None),
            "Midi_solaire": pd.DatetimeIndex(
                sun_events["transit"]
            )
            .tz_convert("Europe/Paris")
            .tz_localize(None),
        }
    )

    result = result.merge(
        sun_table,
        on="Date_solaire",
        how="left",
    )

    # Contrôle secondaire par comparaison directe aux événements.
    result["Dans_intervalle_lever_coucher"] = (
        result["Lever_soleil"].notna()
        & result["Coucher_soleil"].notna()
        & (result["Horodate_milieu"] >= result["Lever_soleil"])
        & (result["Horodate_milieu"] <= result["Coucher_soleil"])
    )

    # Si les événements sont disponibles, ils doivent être cohérents avec
    # l'élévation. La position solaire reste toutefois la référence principale.
    result["Controle_solaire_coherent"] = (
        result["Soleil_leve"]
        == result["Dans_intervalle_lever_coucher"]
    )

    return result


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_pvgis_reference_profile(
    latitude: float,
    longitude: float,
    tilt: float,
    aspect: float,
    peak_power_kwp: float,
    losses_percent: float,
) -> tuple[pd.DataFrame, dict]:
    """
    Récupère un profil horaire PVGIS récent (2020-2023), puis calcule un
    profil de référence moyen par mois, jour et heure.

    PVGIS exprime aspect ainsi :
    0 = sud, -90 = est, 90 = ouest.
    """
    url = "https://re.jrc.ec.europa.eu/api/v5_3/seriescalc"

    params = {
        "lat": latitude,
        "lon": longitude,
        "startyear": 2020,
        "endyear": 2023,
        "pvcalculation": 1,
        "peakpower": peak_power_kwp,
        "loss": losses_percent,
        "pvtechchoice": "crystSi",
        "mountingplace": "free",
        "angle": tilt,
        "aspect": aspect,
        "components": 1,
        "usehorizon": 1,
        "outputformat": "json",
    }

    response = requests.get(
        url,
        params=params,
        timeout=120,
    )
    response.raise_for_status()
    payload = response.json()

    hourly_rows = payload.get("outputs", {}).get("hourly", [])

    if not hourly_rows:
        raise ValueError("PVGIS n'a renvoyé aucune donnée horaire.")

    pvgis = pd.DataFrame(hourly_rows)

    pvgis["Datetime_UTC"] = pd.to_datetime(
        pvgis["time"],
        format="%Y%m%d:%H%M",
        utc=True,
        errors="coerce",
    )
    pvgis = pvgis.dropna(subset=["Datetime_UTC"]).copy()

    pvgis["Datetime_local"] = (
        pvgis["Datetime_UTC"]
        .dt.tz_convert("Europe/Paris")
        .dt.tz_localize(None)
    )

    if "G(i)" in pvgis.columns:
        pvgis["Irradiation_Wm2"] = pd.to_numeric(
            pvgis["G(i)"],
            errors="coerce",
        )
    else:
        component_columns = [
            column
            for column in ["Gb(i)", "Gd(i)", "Gr(i)"]
            if column in pvgis.columns
        ]

        if component_columns:
            pvgis["Irradiation_Wm2"] = (
                pvgis[component_columns]
                .apply(pd.to_numeric, errors="coerce")
                .sum(axis=1)
            )
        else:
            pvgis["Irradiation_Wm2"] = np.nan

    pvgis["Production_PV_kW"] = (
        pd.to_numeric(pvgis.get("P"), errors="coerce") / 1000
    )

    pvgis["Mois"] = pvgis["Datetime_local"].dt.month
    pvgis["Jour_mois"] = pvgis["Datetime_local"].dt.day
    pvgis["Heure"] = pvgis["Datetime_local"].dt.hour

    exact_profile = (
        pvgis.groupby(
            ["Mois", "Jour_mois", "Heure"],
            as_index=False,
        )
        .agg(
            Irradiation_Wm2=("Irradiation_Wm2", "mean"),
            Production_PV_kW=("Production_PV_kW", "mean"),
        )
    )

    metadata = payload.get("inputs", {})
    metadata["source_period"] = "2020-2023"

    return exact_profile, metadata


def merge_pvgis_profile(
    df: pd.DataFrame,
    pvgis_profile: pd.DataFrame,
) -> pd.DataFrame:
    result = df.copy()
    result["Mois_solaire"] = result["Horodate"].dt.month
    result["Jour_mois_solaire"] = result["Horodate"].dt.day
    result["Heure_solaire"] = result["Horodate"].dt.hour

    exact = pvgis_profile.rename(
        columns={
            "Mois": "Mois_solaire",
            "Jour_mois": "Jour_mois_solaire",
            "Heure": "Heure_solaire",
        }
    )

    result = result.merge(
        exact,
        on=[
            "Mois_solaire",
            "Jour_mois_solaire",
            "Heure_solaire",
        ],
        how="left",
    )

    # Repli pour le 29 février ou les rares trous : moyenne mois/heure.
    fallback = (
        pvgis_profile.groupby(
            ["Mois", "Heure"],
            as_index=False,
        )
        .agg(
            Irradiation_fallback=("Irradiation_Wm2", "mean"),
            Production_fallback=("Production_PV_kW", "mean"),
        )
        .rename(
            columns={
                "Mois": "Mois_solaire",
                "Heure": "Heure_solaire",
            }
        )
    )

    result = result.merge(
        fallback,
        on=["Mois_solaire", "Heure_solaire"],
        how="left",
    )

    result["Irradiation_Wm2"] = result[
        "Irradiation_Wm2"
    ].fillna(result["Irradiation_fallback"])

    result["Production_PV_kW"] = result[
        "Production_PV_kW"
    ].fillna(result["Production_fallback"])

    result["Production_PV_kWh"] = (
        result["Production_PV_kW"] * result["Duree_h"]
    )

    result["Autoconsommation_estimee_kWh"] = np.minimum(
        result["Energie_kWh"],
        result["Production_PV_kWh"].fillna(0),
    )

    return result


def build_daily_solar_summary(df: pd.DataFrame) -> pd.DataFrame:
    if "Lever_soleil" not in df.columns:
        return pd.DataFrame()

    daily = (
        df.groupby(df["Horodate"].dt.normalize(), as_index=False)
        .agg(
            Lever_soleil=("Lever_soleil", "first"),
            Coucher_soleil=("Coucher_soleil", "first"),
            Midi_solaire=("Midi_solaire", "first"),
            Irradiation_moyenne_Wm2=("Irradiation_Wm2", "mean"),
            Irradiation_max_Wm2=("Irradiation_Wm2", "max"),
            Production_PV_kWh=("Production_PV_kWh", "sum"),
        )
        .rename(columns={"Horodate": "Date"})
    )

    daily["Duree_jour_h"] = (
        daily["Coucher_soleil"] - daily["Lever_soleil"]
    ).dt.total_seconds() / 3600

    return daily




def clamp_score(value: float) -> float:
    if pd.isna(value):
        return 0.0
    return float(max(0.0, min(100.0, value)))


def score_linear(value: float, low: float, high: float) -> float:
    """Transforme une valeur en score 0-100 entre deux bornes."""
    if pd.isna(value) or high <= low:
        return 0.0
    return clamp_score((value - low) / (high - low) * 100)


def score_status(score: float) -> tuple[str, str]:
    if score >= 80:
        return "Très favorable", "#2E8B57"
    if score >= 65:
        return "Favorable", "#69A84F"
    if score >= 45:
        return "À approfondir", "#E0A800"
    if score >= 25:
        return "Vigilance", "#E67E22"
    return "Peu favorable", "#C0392B"


def metric_status(value: float, thresholds: list[tuple[float, str, str]]) -> tuple[str, str]:
    """Retourne le premier statut dont le seuil minimal est atteint."""
    if pd.isna(value):
        return "Non disponible", "#7B8794"
    for minimum, label, color in thresholds:
        if value >= minimum:
            return label, color
    return "À contrôler", "#C0392B"


def calculate_cma_pv_score(
    production_period_share: float,
    self_consumption_rate: float,
    self_sufficiency_rate: float,
    daily_consumption: pd.Series,
    annual_yield_kwh_per_kwp: float,
) -> dict:
    """
    Indice pédagogique CMA sur 100.

    Pondérations :
    - 30 % correspondance consommation / production PV ;
    - 25 % taux d'autoconsommation ;
    - 20 % taux d'autoproduction ;
    - 15 % régularité de la consommation journalière ;
    - 10 % potentiel solaire local.
    """
    overlap_score = score_linear(production_period_share, 20, 80)
    self_consumption_score = score_linear(self_consumption_rate, 45, 95)
    self_sufficiency_score = score_linear(self_sufficiency_rate, 5, 45)

    daily_clean = pd.to_numeric(daily_consumption, errors="coerce").dropna()
    if daily_clean.empty or daily_clean.mean() <= 0:
        regularity_score = 0.0
        coefficient_variation = np.nan
    else:
        coefficient_variation = daily_clean.std(ddof=0) / daily_clean.mean()
        # CV 0,10 = très régulier ; CV 1,00 = très irrégulier.
        regularity_score = clamp_score(
            (1 - (coefficient_variation - 0.10) / 0.90) * 100
        )

    # Repères volontairement larges pour la France métropolitaine.
    solar_resource_score = score_linear(
        annual_yield_kwh_per_kwp,
        800,
        1400,
    )

    weighted_score = (
        overlap_score * 0.30
        + self_consumption_score * 0.25
        + self_sufficiency_score * 0.20
        + regularity_score * 0.15
        + solar_resource_score * 0.10
    )

    label, color = score_status(weighted_score)

    return {
        "score": round(weighted_score, 1),
        "label": label,
        "color": color,
        "overlap_score": round(overlap_score, 1),
        "self_consumption_score": round(self_consumption_score, 1),
        "self_sufficiency_score": round(self_sufficiency_score, 1),
        "regularity_score": round(regularity_score, 1),
        "solar_resource_score": round(solar_resource_score, 1),
        "coefficient_variation": coefficient_variation,
    }


def build_cma_score_comment(score_data: dict) -> str:
    score = score_data["score"]

    if score >= 80:
        return (
            "Le profil présente une très bonne correspondance entre les usages "
            "électriques et la production solaire simulée. Le projet paraît "
            "particulièrement intéressant à approfondir."
        )
    if score >= 65:
        return (
            "Le profil est globalement favorable au photovoltaïque. Une étude "
            "technique et économique permettra de confirmer la puissance la "
            "plus pertinente."
        )
    if score >= 45:
        return (
            "Le potentiel est réel mais plusieurs paramètres doivent être "
            "approfondis, notamment la puissance installée, les horaires de "
            "consommation et la gestion du surplus."
        )
    if score >= 25:
        return (
            "La correspondance entre les besoins et la production solaire "
            "reste limitée. Un scénario plus modeste ou une adaptation des "
            "usages peut améliorer le projet."
        )
    return (
        "Le profil paraît peu favorable dans le scénario actuellement étudié. "
        "Cela ne signifie pas que le photovoltaïque est impossible, mais qu'un "
        "dimensionnement différent doit être envisagé."
    )


def render_status_pill(label: str, color: str) -> str:
    return (
        f'<span class="status-pill" style="background:{color};">'
        f'{label}</span>'
    )


def render_score_card(score_data: dict) -> None:
    comment = build_cma_score_comment(score_data)
    st.markdown(
        f"""
        <div class="score-card" style="--score-color:{score_data['color']};">
            <div class="score-circle">
                <div class="score-number">{score_data['score']:.0f}</div>
                <div class="score-total">sur 100</div>
            </div>
            <div>
                <div class="score-title">
                    Indice photovoltaïque CMA — {score_data['label']}
                </div>
                <p class="score-text">{comment}</p>
                {render_status_pill(score_data['label'], score_data['color'])}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def time_to_minutes(value) -> int:
    return int(value.hour) * 60 + int(value.minute)


def minutes_in_time_range(
    minute_of_day: int,
    start_minute: int,
    end_minute: int,
) -> bool:
    """
    Vérifie si une minute appartient à une plage horaire.
    Une plage dont la fin est antérieure au début traverse minuit.
    Exemple : 22:00 -> 06:00.
    """
    if start_minute == end_minute:
        return True
    if start_minute < end_minute:
        return start_minute <= minute_of_day < end_minute
    return minute_of_day >= start_minute or minute_of_day < end_minute


def add_tariff_categories(
    df: pd.DataFrame,
    hc_ranges: list[tuple],
) -> pd.DataFrame:
    """
    Classe chaque intervalle Enedis selon :
    - saison haute / hiver : 1er novembre au 31 mars ;
    - saison basse / été : 1er avril au 31 octobre ;
    - heures creuses selon les plages saisies ;
    - heures pleines par complément.

    Le classement utilise le milieu réel de l'intervalle.
    """
    result = df.copy()

    if "Duree_h" not in result.columns:
        result["Duree_h"] = 1.0

    result["Horodate_tarif"] = (
        result["Horodate"]
        - pd.to_timedelta(result["Duree_h"] / 2, unit="h")
    )

    hc_ranges_minutes = [
        (time_to_minutes(start), time_to_minutes(end))
        for start, end in hc_ranges
    ]

    minute_of_day = (
        result["Horodate_tarif"].dt.hour * 60
        + result["Horodate_tarif"].dt.minute
    )

    hc_mask = pd.Series(False, index=result.index)

    for start_minute, end_minute in hc_ranges_minutes:
        if start_minute == end_minute:
            hc_mask = pd.Series(True, index=result.index)
        elif start_minute < end_minute:
            hc_mask |= (
                (minute_of_day >= start_minute)
                & (minute_of_day < end_minute)
            )
        else:
            hc_mask |= (
                (minute_of_day >= start_minute)
                | (minute_of_day < end_minute)
            )

    month = result["Horodate_tarif"].dt.month
    winter_mask = month.isin([11, 12, 1, 2, 3])

    result["Saison_tarifaire"] = np.where(
        winter_mask,
        "Hiver / saison haute",
        "Été / saison basse",
    )
    result["Plage_tarifaire"] = np.where(
        hc_mask,
        "Heures creuses",
        "Heures pleines",
    )

    result["Categorie_tarifaire"] = np.select(
        [
            winter_mask & ~hc_mask,
            winter_mask & hc_mask,
            ~winter_mask & ~hc_mask,
            ~winter_mask & hc_mask,
        ],
        [
            "HP hiver",
            "HC hiver",
            "HP été",
            "HC été",
        ],
        default="Non classé",
    )

    return result


def build_tariff_summary(df: pd.DataFrame) -> pd.DataFrame:
    order = ["HP hiver", "HC hiver", "HP été", "HC été"]

    summary = (
        df.groupby("Categorie_tarifaire", as_index=False)
        .agg(
            Consommation_kWh=("Energie_kWh", "sum"),
            Nombre_intervalles=("Energie_kWh", "size"),
        )
    )

    summary = (
        summary.set_index("Categorie_tarifaire")
        .reindex(order, fill_value=0)
        .reset_index()
    )

    total = summary["Consommation_kWh"].sum()
    summary["Part_pourcent"] = np.where(
        total > 0,
        summary["Consommation_kWh"] / total * 100,
        0,
    )

    return summary


def calculate_tariff_optimization_score(
    tariff_summary: pd.DataFrame,
    daily_consumption: pd.Series,
    coverage_ratio: float,
) -> dict:
    """
    Indice de potentiel d'optimisation tarifaire.

    Un score élevé ne signifie pas que le contrat est mauvais.
    Il indique qu'une analyse plus poussée peut être utile.

    Pondérations :
    - 50 % : part consommée en heures pleines ;
    - 25 % : déséquilibre saison haute / saison basse ;
    - 15 % : régularité des consommations ;
    - 10 % : qualité / complétude de la période analysée.
    """
    values = tariff_summary.set_index("Categorie_tarifaire")[
        "Consommation_kWh"
    ]

    hp_total = float(
        values.get("HP hiver", 0)
        + values.get("HP été", 0)
    )
    hc_total = float(
        values.get("HC hiver", 0)
        + values.get("HC été", 0)
    )
    winter_total = float(
        values.get("HP hiver", 0)
        + values.get("HC hiver", 0)
    )
    summer_total = float(
        values.get("HP été", 0)
        + values.get("HC été", 0)
    )
    total = hp_total + hc_total

    hp_share = hp_total / total * 100 if total else 0
    hc_share = hc_total / total * 100 if total else 0

    # Plus la part HP est forte, plus il existe potentiellement des usages
    # à examiner pour un déplacement en HC.
    hp_opportunity_score = score_linear(hp_share, 45, 90)

    # Comparaison corrigée par le nombre de mois des deux saisons.
    winter_monthly_average = winter_total / 5
    summer_monthly_average = summer_total / 7

    if max(winter_monthly_average, summer_monthly_average) > 0:
        seasonal_gap = (
            abs(winter_monthly_average - summer_monthly_average)
            / max(winter_monthly_average, summer_monthly_average)
            * 100
        )
    else:
        seasonal_gap = 0

    seasonality_score = score_linear(seasonal_gap, 10, 70)

    daily_clean = pd.to_numeric(
        daily_consumption,
        errors="coerce",
    ).dropna()

    if daily_clean.empty or daily_clean.mean() <= 0:
        regularity_score = 0
    else:
        cv = daily_clean.std(ddof=0) / daily_clean.mean()
        # Pour un potentiel d'optimisation, une forte variabilité crée plus
        # de matière à analyser. Le score n'est pas un jugement de qualité.
        regularity_score = score_linear(cv, 0.15, 0.90)

    coverage_score = clamp_score(coverage_ratio * 100)

    score = (
        hp_opportunity_score * 0.50
        + seasonality_score * 0.25
        + regularity_score * 0.15
        + coverage_score * 0.10
    )

    if score >= 75:
        label = "Potentiel d'analyse élevé"
        color = "#C0392B"
    elif score >= 55:
        label = "Potentiel d'analyse significatif"
        color = "#E67E22"
    elif score >= 35:
        label = "Potentiel d'analyse modéré"
        color = "#E0A800"
    else:
        label = "Potentiel d'analyse limité"
        color = "#2E8B57"

    return {
        "score": round(score, 1),
        "label": label,
        "color": color,
        "hp_share": round(hp_share, 1),
        "hc_share": round(hc_share, 1),
        "seasonal_gap": round(seasonal_gap, 1),
        "hp_opportunity_score": round(hp_opportunity_score, 1),
        "seasonality_score": round(seasonality_score, 1),
        "variability_score": round(regularity_score, 1),
        "coverage_score": round(coverage_score, 1),
        "winter_total": winter_total,
        "summer_total": summer_total,
    }


def build_tariff_commentary(score_data: dict) -> str:
    hp_share = score_data["hp_share"]
    seasonal_gap = score_data["seasonal_gap"]

    if hp_share >= 75:
        hp_text = (
            "La consommation est très majoritairement située en heures "
            "pleines. Il peut être utile d'identifier les usages décalables, "
            "sans perturber l'activité."
        )
    elif hp_share >= 55:
        hp_text = (
            "La majorité de la consommation a lieu en heures pleines. "
            "Certains usages non critiques peuvent éventuellement être "
            "examinés pour un déplacement en heures creuses."
        )
    else:
        hp_text = (
            "Une part importante de la consommation est déjà réalisée en "
            "heures creuses. Le potentiel de déplacement supplémentaire "
            "semble plus limité."
        )

    if seasonal_gap >= 45:
        season_text = (
            "La consommation moyenne mensuelle diffère fortement entre les "
            "saisons haute et basse. Il convient d'en rechercher les causes "
            "possibles : chauffage, process, activité saisonnière ou occupation."
        )
    elif seasonal_gap >= 20:
        season_text = (
            "Une saisonnalité modérée est visible entre hiver et été. "
            "Elle mérite d'être mise en regard des usages de l'entreprise."
        )
    else:
        season_text = (
            "La consommation est relativement homogène entre les deux saisons."
        )

    return hp_text + " " + season_text



# ============================================================
# MOTEUR FINANCIER PHOTOVOLTAÏQUE
# ============================================================

PV_FIXING_COSTS_EUR_WC = {
    "Surimposition": 0.30,
    "Bac acier + panneaux": 0.10,
    "Bac acier isolé + panneaux": 0.10,
    "Lesté en toiture terrasse": 0.25,
    "Soudé sur toiture terrasse": 0.20,
    "Intégré au bâti": 0.55,
    "Installation au sol": 0.35,
    "Ombrière": 0.70,
}

ROOF_RENOVATION_COSTS_EUR_M2 = {
    "Bac acier isolé": 140.0,
    "Bac acier": 90.0,
    "Fibrociment": 90.0,
    "Tuile": 80.0,
    "Ardoise": 150.0,
    "Toiture terrasse": 150.0,
}

ASBESTOS_REMOVAL_EUR_M2 = 60.0


def installation_cost_eur_wc(peak_power_kwp: float) -> float:
    """Coût modules + onduleur + câblage + pose, hors fixation."""
    if peak_power_kwp <= 9:
        return 1.30
    if peak_power_kwp <= 36:
        return 0.75
    if peak_power_kwp <= 100:
        return 0.65
    return 0.45


def turpe_annual_cost(peak_power_kwp: float) -> float:
    if peak_power_kwp <= 36:
        return 9.24
    if peak_power_kwp <= 250:
        return 136.08
    return 272.40


def ifer_annual_cost(
    peak_power_kwp: float,
    rate_eur_kwp: float = 3.542,
) -> float:
    return (
        peak_power_kwp * rate_eur_kwp
        if peak_power_kwp > 100
        else 0.0
    )


def calculate_connection_cost(
    peak_power_kwp: float,
    connection_mode: str,
    public_extension_length_m: float,
    private_trench_length_m: float,
    apply_enedis_reduction: bool,
    include_private_hta_post: bool,
    include_decoupling_cell: bool,
) -> dict:
    """
    Estimation indicative du raccordement en vente de surplus.

    La réfaction Enedis est appliquée uniquement aux ouvrages publics
    calculés dans cette fonction. Les ouvrages privés restent à la charge
    du porteur de projet.
    """
    if peak_power_kwp <= 36:
        public_gross = 0.0
    elif connection_mode == "Branchement BT avec extension":
        public_gross = (
            1000.0
            + 1000.0
            + 50.0 * public_extension_length_m
        )
    elif connection_mode == "Branchement complet C4":
        public_gross = 1700.0
    elif connection_mode == "Création poste BT-HTA":
        public_gross = (
            8800.0
            + 2200.0
            + 50.0 * public_extension_length_m
        )
    else:
        public_gross = 0.0

    reduction = (
        public_gross * 0.60
        if apply_enedis_reduction and peak_power_kwp > 36
        else 0.0
    )
    public_net = max(public_gross - reduction, 0.0)

    private_trench = 125.0 * private_trench_length_m
    private_post = 55000.0 if include_private_hta_post else 0.0
    decoupling_cell = 15000.0 if include_decoupling_cell else 0.0

    private_total = private_trench + private_post + decoupling_cell
    total = public_net + private_total

    return {
        "public_gross": public_gross,
        "enedis_reduction": reduction,
        "public_net": public_net,
        "private_trench": private_trench,
        "private_post": private_post,
        "decoupling_cell": decoupling_cell,
        "private_total": private_total,
        "total": total,
    }


def calculate_investment_costs(
    peak_power_kwp: float,
    fixing_type: str,
    erp_icpe_surcharge: bool,
    structural_study_cost: float,
    roof_renovation_enabled: bool,
    roof_type: str,
    roof_area_m2: float,
    asbestos_removal_enabled: bool,
    connection_data: dict,
    other_investment_costs: float,
    grant_amount: float,
) -> dict:
    power_wc = peak_power_kwp * 1000.0

    fixing_rate = PV_FIXING_COSTS_EUR_WC[fixing_type]
    equipment_rate = installation_cost_eur_wc(peak_power_kwp)
    erp_rate = 0.10 if erp_icpe_surcharge else 0.0

    fixing_cost = fixing_rate * power_wc
    equipment_cost = equipment_rate * power_wc
    erp_surcharge_cost = erp_rate * power_wc

    roof_cost = (
        ROOF_RENOVATION_COSTS_EUR_M2[roof_type] * roof_area_m2
        if roof_renovation_enabled
        else 0.0
    )
    asbestos_cost = (
        ASBESTOS_REMOVAL_EUR_M2 * roof_area_m2
        if asbestos_removal_enabled
        else 0.0
    )

    gross_total = (
        fixing_cost
        + equipment_cost
        + erp_surcharge_cost
        + structural_study_cost
        + roof_cost
        + asbestos_cost
        + connection_data["total"]
        + other_investment_costs
    )

    net_total = max(gross_total - grant_amount, 0.0)

    return {
        "power_wc": power_wc,
        "fixing_rate": fixing_rate,
        "equipment_rate": equipment_rate,
        "fixing_cost": fixing_cost,
        "equipment_cost": equipment_cost,
        "erp_surcharge_cost": erp_surcharge_cost,
        "structural_study_cost": structural_study_cost,
        "roof_cost": roof_cost,
        "asbestos_cost": asbestos_cost,
        "connection_cost": connection_data["total"],
        "other_investment_costs": other_investment_costs,
        "gross_total": gross_total,
        "grant_amount": grant_amount,
        "net_total": net_total,
    }


def tariff_price_map(
    tariff_type: str,
    unique_price: float,
    hp_price: float,
    hc_price: float,
    hp_winter_price: float,
    hc_winter_price: float,
    hp_summer_price: float,
    hc_summer_price: float,
) -> dict:
    if tariff_type == "Tarif unique":
        return {
            "HP hiver": unique_price,
            "HC hiver": unique_price,
            "HP été": unique_price,
            "HC été": unique_price,
        }

    if tariff_type == "HP / HC":
        return {
            "HP hiver": hp_price,
            "HC hiver": hc_price,
            "HP été": hp_price,
            "HC été": hc_price,
        }

    return {
        "HP hiver": hp_winter_price,
        "HC hiver": hc_winter_price,
        "HP été": hp_summer_price,
        "HC été": hc_summer_price,
    }


def calculate_energy_value(
    df: pd.DataFrame,
    price_map: dict,
    surplus_sale_price: float,
    annual_subscription: float,
    analysis_years: float,
) -> dict:
    result = df.copy()
    result["Prix_achat_EUR_kWh"] = (
        result["Categorie_tarifaire"]
        .map(price_map)
        .fillna(0.0)
    )

    result["Cout_electricite_avant_PV_EUR"] = (
        result["Energie_kWh"]
        * result["Prix_achat_EUR_kWh"]
    )

    autoconsumed = (
        result["Autoconsommation_estimee_kWh"].fillna(0.0)
        if "Autoconsommation_estimee_kWh" in result.columns
        else pd.Series(0.0, index=result.index)
    )

    pv_production = (
        result["Production_PV_kWh"].fillna(0.0)
        if "Production_PV_kWh" in result.columns
        else pd.Series(0.0, index=result.index)
    )

    result["Economie_autoconsommation_EUR"] = (
        autoconsumed * result["Prix_achat_EUR_kWh"]
    )
    result["Surplus_PV_kWh"] = (
        pv_production - autoconsumed
    ).clip(lower=0.0)
    result["Revenu_surplus_EUR"] = (
        result["Surplus_PV_kWh"] * surplus_sale_price
    )

    safe_years = max(float(analysis_years), 1 / 365.25)
    annual_factor = 1.0 / safe_years

    annual_energy_bill = (
        result["Cout_electricite_avant_PV_EUR"].sum()
        * annual_factor
        + annual_subscription
    )
    annual_self_consumption_saving = (
        result["Economie_autoconsommation_EUR"].sum()
        * annual_factor
    )
    annual_surplus_revenue = (
        result["Revenu_surplus_EUR"].sum()
        * annual_factor
    )

    return {
        "detail": result,
        "annual_factor": annual_factor,
        "annual_energy_bill": annual_energy_bill,
        "annual_self_consumption_saving": annual_self_consumption_saving,
        "annual_surplus_revenue": annual_surplus_revenue,
    }


def calculate_annual_operating_costs(
    peak_power_kwp: float,
    investment_gross: float,
    insurance_rate_percent: float,
    maintenance_eur_kwp: float,
    inverter_provision_eur_kwp: float,
    ifer_rate_eur_kwp: float,
    other_annual_costs: float,
) -> dict:
    insurance = (
        investment_gross
        * insurance_rate_percent
        / 100.0
    )
    maintenance = maintenance_eur_kwp * peak_power_kwp
    inverter_provision = (
        inverter_provision_eur_kwp * peak_power_kwp
    )
    turpe = turpe_annual_cost(peak_power_kwp)
    ifer = ifer_annual_cost(
        peak_power_kwp,
        ifer_rate_eur_kwp,
    )

    total = (
        insurance
        + maintenance
        + inverter_provision
        + turpe
        + ifer
        + other_annual_costs
    )

    return {
        "insurance": insurance,
        "maintenance": maintenance,
        "inverter_provision": inverter_provision,
        "turpe": turpe,
        "ifer": ifer,
        "other": other_annual_costs,
        "total": total,
    }


def npv_from_cashflows(
    cashflows: list[float],
    discount_rate: float,
) -> float:
    return sum(
        cashflow / ((1.0 + discount_rate) ** year)
        for year, cashflow in enumerate(cashflows)
    )


def irr_from_cashflows(
    cashflows: list[float],
    lower: float = -0.99,
    upper: float = 5.0,
) -> float:
    """TRI annuel calculé par dichotomie, sans dépendance supplémentaire."""
    def value(rate: float) -> float:
        return npv_from_cashflows(cashflows, rate)

    low_value = value(lower)
    high_value = value(upper)

    if low_value * high_value > 0:
        return np.nan

    for _ in range(150):
        middle = (lower + upper) / 2.0
        middle_value = value(middle)

        if abs(middle_value) < 1e-8:
            return middle

        if low_value * middle_value <= 0:
            upper = middle
            high_value = middle_value
        else:
            lower = middle
            low_value = middle_value

    return (lower + upper) / 2.0


def build_financial_projection(
    net_investment: float,
    annual_self_consumption_saving: float,
    annual_surplus_revenue: float,
    annual_operating_cost: float,
    horizon_years: int,
    electricity_price_increase_percent: float,
    surplus_price_increase_percent: float,
    production_degradation_percent: float,
    operating_cost_increase_percent: float,
    discount_rate_percent: float,
) -> dict:
    rows = []
    cumulative = -net_investment
    discounted_cumulative = -net_investment
    cashflows = [-net_investment]
    payback_year = np.nan

    for year in range(1, horizon_years + 1):
        production_factor = (
            1.0 - production_degradation_percent / 100.0
        ) ** (year - 1)
        electricity_factor = (
            1.0 + electricity_price_increase_percent / 100.0
        ) ** (year - 1)
        surplus_factor = (
            1.0 + surplus_price_increase_percent / 100.0
        ) ** (year - 1)
        opex_factor = (
            1.0 + operating_cost_increase_percent / 100.0
        ) ** (year - 1)

        self_consumption_saving = (
            annual_self_consumption_saving
            * production_factor
            * electricity_factor
        )
        surplus_revenue = (
            annual_surplus_revenue
            * production_factor
            * surplus_factor
        )
        operating_cost = (
            annual_operating_cost * opex_factor
        )
        net_cashflow = (
            self_consumption_saving
            + surplus_revenue
            - operating_cost
        )

        discount_factor = (
            1.0 + discount_rate_percent / 100.0
        ) ** year
        discounted_cashflow = net_cashflow / discount_factor

        previous_cumulative = cumulative
        cumulative += net_cashflow
        discounted_cumulative += discounted_cashflow
        cashflows.append(net_cashflow)

        if (
            pd.isna(payback_year)
            and cumulative >= 0
            and net_cashflow > 0
        ):
            fraction = (
                -previous_cumulative / net_cashflow
                if previous_cumulative < 0
                else 0.0
            )
            payback_year = year - 1 + fraction

        rows.append(
            {
                "Année": year,
                "Économie autoconsommation (€)": self_consumption_saving,
                "Revenu surplus (€)": surplus_revenue,
                "Charges annuelles (€)": operating_cost,
                "Flux net (€)": net_cashflow,
                "Flux actualisé (€)": discounted_cashflow,
                "Cumul net (€)": cumulative,
                "Cumul actualisé (€)": discounted_cumulative,
            }
        )

    discount_rate = discount_rate_percent / 100.0
    npv = npv_from_cashflows(cashflows, discount_rate)
    irr = irr_from_cashflows(cashflows)

    return {
        "table": pd.DataFrame(rows),
        "cashflows": cashflows,
        "payback_year": payback_year,
        "npv": npv,
        "irr": irr,
        "total_net_gain": cumulative,
        "annual_net_gain_year_1": (
            rows[0]["Flux net (€)"] if rows else 0.0
        ),
    }



# ============================================================
# ASSISTANT MÉTIER CMA — RÈGLES EXPLICITES, SANS IA EXTERNE
# ============================================================

def build_cma_business_assistant(
    daylight_share: float,
    production_period_share: float,
    self_consumption_rate: float,
    self_sufficiency_rate: float,
    pv_surplus_kwh: float,
    pvgis_production_kwh: float,
    cma_score_data: dict,
    tariff_score_data: dict,
    investment_data: dict,
    operating_cost_data: dict,
    financial_projection: dict,
    roof_renovation_enabled: bool,
    asbestos_removal_enabled: bool,
    connection_data: dict,
    coverage_ratio: float,
) -> dict:
    """Produit une synthèse métier traçable à partir de règles explicites."""
    strengths = []
    vigilance = []
    next_steps = []
    conclusion_parts = []

    # Profil de consommation
    if production_period_share >= 65:
        strengths.append(
            "Une part importante de la consommation intervient pendant "
            "les périodes de production photovoltaïque."
        )
    elif production_period_share >= 45:
        strengths.append(
            "La correspondance entre les consommations et la production "
            "solaire est globalement favorable."
        )
    else:
        vigilance.append(
            "Une part limitée des consommations coïncide avec la production "
            "solaire ; le dimensionnement devra rester prudent."
        )

    # Autoconsommation
    if self_consumption_rate >= 85:
        strengths.append(
            "Le taux d'autoconsommation simulé est très élevé, ce qui limite "
            "le surplus injecté sur le réseau."
        )
    elif self_consumption_rate >= 65:
        strengths.append(
            "Le taux d'autoconsommation simulé est satisfaisant."
        )
    else:
        vigilance.append(
            "Le surplus photovoltaïque paraît important ; une puissance "
            "inférieure ou un déplacement d'usages peut être étudié."
        )

    # Autoproduction
    if self_sufficiency_rate >= 35:
        strengths.append(
            "Le projet pourrait couvrir une part significative des besoins "
            "électriques du site."
        )
    elif self_sufficiency_rate < 15:
        vigilance.append(
            "Le taux d'autoproduction reste faible : le réseau demeurera la "
            "source principale d'électricité."
        )

    # Surplus
    surplus_ratio = (
        pv_surplus_kwh / pvgis_production_kwh * 100
        if pvgis_production_kwh and not pd.isna(pvgis_production_kwh)
        else 0
    )
    if surplus_ratio > 35:
        vigilance.append(
            f"Environ {format_fr(surplus_ratio, 0)} % de la production "
            "simulée serait injectée en surplus."
        )

    # Finance
    payback = financial_projection.get("payback_year", np.nan)
    npv = financial_projection.get("npv", np.nan)
    irr = financial_projection.get("irr", np.nan)

    if not pd.isna(payback):
        if payback <= 8:
            strengths.append(
                f"Le temps de retour simple est favorable, estimé à "
                f"{format_fr(payback, 1)} ans."
            )
        elif payback <= 12:
            conclusion_parts.append(
                f"Le temps de retour simple est estimé à "
                f"{format_fr(payback, 1)} ans."
            )
        else:
            vigilance.append(
                f"Le temps de retour simple est relativement long "
                f"({format_fr(payback, 1)} ans)."
            )
    else:
        vigilance.append(
            "L'investissement n'est pas amorti sur l'horizon de projection."
        )

    if not pd.isna(npv) and npv < 0:
        vigilance.append(
            "La valeur actuelle nette est négative avec les hypothèses "
            "retenues."
        )
    elif not pd.isna(npv) and npv > 0:
        strengths.append(
            "La valeur actuelle nette est positive avec les hypothèses "
            "retenues."
        )

    if not pd.isna(irr):
        conclusion_parts.append(
            f"Le TRI indicatif ressort à {format_fr(irr * 100, 1)} %."
        )

    # Données et contraintes
    if coverage_ratio < 0.90:
        vigilance.append(
            "La période de données couvre moins d'une année complète ; "
            "l'annualisation doit être interprétée avec prudence."
        )

    if roof_renovation_enabled:
        vigilance.append(
            "Le scénario intègre une rénovation de couverture, poste à "
            "confirmer par devis."
        )
        next_steps.append(
            "Faire établir un diagnostic de toiture et un devis de couverture."
        )

    if asbestos_removal_enabled:
        vigilance.append(
            "Un coût indicatif de désamiantage est intégré au scénario."
        )
        next_steps.append(
            "Faire confirmer la présence d'amiante et le protocole de retrait."
        )

    if connection_data.get("total", 0) > 0:
        next_steps.append(
            "Demander une proposition de raccordement afin de sécuriser le "
            "coût réel et les délais."
        )

    if investment_data.get("structural_study_cost", 0) > 0:
        next_steps.append(
            "Faire valider la capacité portante de la charpente et de la toiture."
        )

    next_steps.extend(
        [
            "Comparer plusieurs puissances et plusieurs devis d'installateurs qualifiés.",
            "Vérifier les aides, la fiscalité et le mode de financement.",
            "Actualiser les prix d'achat et de vente de l'électricité avant décision.",
        ]
    )

    score = float(cma_score_data.get("score", 0))
    if score >= 80 and (pd.isna(payback) or payback <= 10):
        headline = "Profil très favorable à approfondir"
        status = "Très favorable"
        color = "#2E8B57"
    elif score >= 65:
        headline = "Projet globalement favorable"
        status = "Favorable"
        color = "#69A84F"
    elif score >= 45:
        headline = "Projet intéressant sous conditions"
        status = "À approfondir"
        color = "#E0A800"
    else:
        headline = "Projet nécessitant des ajustements"
        status = "Vigilance"
        color = "#E67E22"

    conclusion = (
        f"{headline}. "
        + " ".join(conclusion_parts)
        + " Les résultats restent indicatifs et doivent être confirmés "
        "par une étude technique, des devis et une analyse financière complète."
    )

    # Déduplication tout en conservant l'ordre.
    strengths = list(dict.fromkeys(strengths))
    vigilance = list(dict.fromkeys(vigilance))
    next_steps = list(dict.fromkeys(next_steps))

    return {
        "headline": headline,
        "status": status,
        "color": color,
        "conclusion": conclusion,
        "strengths": strengths,
        "vigilance": vigilance,
        "next_steps": next_steps,
        "pv_score": cma_score_data.get("score", 0),
        "tariff_score": tariff_score_data.get("score", 0),
    }


def assistant_html_list(items: list[str], icon: str) -> str:
    if not items:
        return "<p>Aucun élément particulier identifié.</p>"
    return "".join(
        f"<div style='margin:7px 0'>{icon} {item}</div>"
        for item in items
    )


def safe_pdf_text(value) -> str:
    if value is None:
        return ""
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def figure_to_png_bytes(fig, width: int = 1100, height: int = 520) -> BytesIO:
    image_bytes = fig.to_image(
        format="png",
        width=width,
        height=height,
        scale=1.5,
    )
    return BytesIO(image_bytes)


def build_automatic_commentary(
    daylight_share: float,
    production_period_share: float,
    self_consumption_rate: float,
    self_sufficiency_rate: float,
) -> dict:
    if pd.isna(production_period_share):
        profile = (
            "L'adresse ou les données PVGIS ne sont pas disponibles. "
            "L'analyse de correspondance solaire n'a pas pu être finalisée."
        )
    elif production_period_share >= 70:
        profile = (
            "Une part très importante de la consommation intervient pendant "
            "les périodes de production photovoltaïque. Le profil paraît "
            "particulièrement favorable à un projet orienté vers "
            "l'autoconsommation."
        )
    elif production_period_share >= 50:
        profile = (
            "Une part significative de la consommation intervient pendant "
            "les périodes de production photovoltaïque. Le profil semble "
            "favorable, sous réserve des vérifications techniques et "
            "économiques complémentaires."
        )
    elif production_period_share >= 30:
        profile = (
            "La consommation correspond partiellement aux périodes de "
            "production solaire. Le projet mérite d'être approfondi afin "
            "d'identifier un dimensionnement adapté."
        )
    else:
        profile = (
            "Une part limitée de la consommation intervient pendant les "
            "périodes de production solaire. Un dimensionnement prudent et "
            "une étude détaillée sont recommandés."
        )

    if pd.isna(self_consumption_rate):
        self_consumption = "Le taux d'autoconsommation n'a pas pu être estimé."
    elif self_consumption_rate >= 80:
        self_consumption = (
            "La majorité de l'électricité produite pourrait être consommée "
            "directement sur le site, ce qui limite le surplus potentiel."
        )
    elif self_consumption_rate >= 55:
        self_consumption = (
            "Une part importante de la production pourrait être utilisée "
            "directement par l'entreprise. Un surplus resterait toutefois "
            "possible à certaines heures."
        )
    else:
        self_consumption = (
            "Une part notable de la production pourrait ne pas être consommée "
            "immédiatement. La puissance étudiée devra être comparée à des "
            "scénarios plus modestes ou à une solution de valorisation du surplus."
        )

    if pd.isna(self_sufficiency_rate):
        self_sufficiency = "Le taux d'autoproduction n'a pas pu être estimé."
    elif self_sufficiency_rate >= 40:
        self_sufficiency = (
            "La simulation indique qu'une part importante des besoins "
            "électriques pourrait être couverte par la production solaire."
        )
    elif self_sufficiency_rate >= 20:
        self_sufficiency = (
            "La production solaire pourrait couvrir une part utile des besoins "
            "électriques, sans rendre l'entreprise indépendante du réseau."
        )
    else:
        self_sufficiency = (
            "La production simulée couvrirait une part limitée des besoins. "
            "Le réseau resterait la principale source d'électricité."
        )

    return {
        "profile": profile,
        "self_consumption": self_consumption,
        "self_sufficiency": self_sufficiency,
    }


def create_cma_pdf_report(
    company_name: str,
    company_siret: str,
    advisor_name: str,
    diagnostic_date,
    address_label: str,
    latitude,
    longitude,
    source_filename: str,
    period_start,
    period_end,
    source_unit: str,
    time_step,
    total_kwh: float,
    average_daily_kwh: float,
    maximum_power_kw: float,
    daylight_share: float,
    production_period_share: float,
    pv_peak_kwp: float,
    pv_tilt: float,
    orientation_label: str,
    pv_losses: float,
    pvgis_production_kwh: float,
    self_consumed_kwh: float,
    self_consumption_rate: float,
    self_sufficiency_rate: float,
    cma_score_data: dict,
    annual_yield_kwh_per_kwp: float,
    tariff_summary_df: pd.DataFrame,
    tariff_score_data: dict,
    hc_ranges: list[tuple],
    investment_data: dict,
    connection_data: dict,
    operating_cost_data: dict,
    energy_value_data: dict,
    financial_projection: dict,
    business_assistant: dict,
    financial_horizon_years: int,
    electricity_tariff_type: str,
    surplus_sale_price_eur_kwh: float,
    electricity_price_increase_percent: float,
    production_degradation_percent: float,
    discount_rate_percent: float,
    monthly_df: pd.DataFrame,
    weekday_hour_matrix: pd.DataFrame,
    hourly_df: pd.DataFrame,
    filtered_df: pd.DataFrame,
    logo_path: Path | None,
) -> bytes:
    output = BytesIO()

    page_width, page_height = A4
    cma_blue = colors.HexColor("#17365D")
    cma_red = colors.HexColor("#E53935")
    light_blue = colors.HexColor("#EAF1F8")
    light_grey = colors.HexColor("#F3F5F7")
    dark_text = colors.HexColor("#202735")

    def draw_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(cma_blue)
        canvas.rect(0, page_height - 1.15 * cm, page_width, 1.15 * cm, fill=1, stroke=0)
        canvas.setFillColor(cma_red)
        canvas.rect(page_width - 4.2 * cm, page_height - 1.15 * cm, 4.2 * cm, 1.15 * cm, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor("#697589"))
        canvas.setFont("Helvetica", 8)
        canvas.drawString(1.6 * cm, 0.8 * cm, "Pré-diagnostic photovoltaïque - CMA Nouvelle-Aquitaine")
        canvas.drawRightString(page_width - 1.6 * cm, 0.8 * cm, f"Page {doc.page}")
        canvas.restoreState()

    frame = Frame(
        1.55 * cm,
        1.35 * cm,
        page_width - 3.1 * cm,
        page_height - 2.9 * cm,
        id="normal",
    )
    template = PageTemplate(id="cma", frames=[frame], onPage=draw_page)
    doc = BaseDocTemplate(
        output,
        pagesize=A4,
        leftMargin=1.55 * cm,
        rightMargin=1.55 * cm,
        topMargin=1.55 * cm,
        bottomMargin=1.35 * cm,
        title="Pré-diagnostic photovoltaïque CMA",
        author="CMA Nouvelle-Aquitaine",
    )
    doc.addPageTemplates([template])

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CMA_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=25,
            leading=29,
            textColor=cma_blue,
            alignment=TA_LEFT,
            spaceAfter=12,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_Subtitle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=12,
            leading=17,
            textColor=dark_text,
            spaceAfter=12,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_H1",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=17,
            leading=21,
            textColor=cma_blue,
            spaceBefore=8,
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_H2",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=16,
            textColor=cma_red,
            spaceBefore=8,
            spaceAfter=7,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.6,
            leading=14,
            textColor=dark_text,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=11,
            textColor=colors.HexColor("#5E6878"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_KPI_Value",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=15,
            leading=18,
            textColor=cma_blue,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CMA_KPI_Label",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=10,
            textColor=colors.HexColor("#667287"),
            alignment=TA_CENTER,
        )
    )

    story = []

    # Couverture
    story.append(Spacer(1, 0.9 * cm))
    if logo_path and logo_path.exists():
        logo = Image(str(logo_path), width=6.0 * cm, height=2.1 * cm, kind="proportional")
        story.append(logo)
        story.append(Spacer(1, 0.45 * cm))

    story.append(Paragraph("Pré-diagnostic photovoltaïque", styles["CMA_Title"]))
    story.append(
        Paragraph(
            "Analyse pédagogique des consommations électriques et simulation "
            "d'une installation photovoltaïque en autoconsommation.",
            styles["CMA_Subtitle"],
        )
    )
    story.append(Spacer(1, 0.45 * cm))

    cover_data = [
        ["Entreprise", safe_pdf_text(company_name or "Non renseignée")],
        ["Adresse", safe_pdf_text(address_label or "Non renseignée")],
        ["SIRET", safe_pdf_text(company_siret or "Non renseigné")],
        ["Conseiller CMA", safe_pdf_text(advisor_name or "Non renseigné")],
        ["Date du diagnostic", diagnostic_date.strftime("%d/%m/%Y")],
        ["Période analysée", f"{period_start:%d/%m/%Y} au {period_end:%d/%m/%Y}"],
    ]
    cover_table = Table(cover_data, colWidths=[4.6 * cm, 11.2 * cm])
    cover_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), light_blue),
                ("TEXTCOLOR", (0, 0), (0, -1), cma_blue),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(cover_table)
    story.append(Spacer(1, 1.0 * cm))
    story.append(
        Paragraph(
            "<b>Important :</b> ce document constitue un pré-diagnostic. "
            "Il ne remplace pas une étude de faisabilité technique, structurelle, "
            "réglementaire et économique réalisée par des professionnels qualifiés.",
            styles["CMA_Body"],
        )
    )
    story.append(PageBreak())

    # Synthèse
    story.append(Paragraph("1. Synthèse de l'analyse", styles["CMA_H1"]))

    score_color = colors.HexColor(cma_score_data.get("color", "#17365D"))
    score_table = Table(
        [
            [
                Paragraph(
                    f"<b>{cma_score_data.get('score', 0):.0f}/100</b>",
                    styles["CMA_KPI_Value"],
                ),
                Paragraph(
                    f"<b>Indice photovoltaïque CMA — "
                    f"{safe_pdf_text(cma_score_data.get('label', 'Non calculé'))}</b><br/>"
                    f"{safe_pdf_text(build_cma_score_comment(cma_score_data))}",
                    styles["CMA_Body"],
                ),
            ]
        ],
        colWidths=[3.3 * cm, 12.5 * cm],
    )
    score_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#F3F6F9")),
                ("BOX", (0, 0), (-1, -1), 1.2, score_color),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    story.append(score_table)
    story.append(Spacer(1, 0.4 * cm))

    kpi_data = [
        [
            Paragraph(f"{format_fr(total_kwh, 0)} kWh", styles["CMA_KPI_Value"]),
            Paragraph(f"{format_fr(average_daily_kwh, 1)} kWh", styles["CMA_KPI_Value"]),
            Paragraph(f"{format_fr(maximum_power_kw, 1)} kW", styles["CMA_KPI_Value"]),
        ],
        [
            Paragraph("Consommation totale", styles["CMA_KPI_Label"]),
            Paragraph("Moyenne par jour", styles["CMA_KPI_Label"]),
            Paragraph("Pic de puissance", styles["CMA_KPI_Label"]),
        ],
        [
            Paragraph(
                f"{format_fr(production_period_share, 1)} %"
                if not pd.isna(production_period_share)
                else "N/D",
                styles["CMA_KPI_Value"],
            ),
            Paragraph(
                f"{format_fr(pvgis_production_kwh, 0)} kWh"
                if not pd.isna(pvgis_production_kwh)
                else "N/D",
                styles["CMA_KPI_Value"],
            ),
            Paragraph(
                f"{format_fr(self_consumption_rate, 1)} %"
                if not pd.isna(self_consumption_rate)
                else "N/D",
                styles["CMA_KPI_Value"],
            ),
        ],
        [
            Paragraph("Conso. pendant la production PV", styles["CMA_KPI_Label"]),
            Paragraph("Production PV estimée", styles["CMA_KPI_Label"]),
            Paragraph("Taux d'autoconsommation", styles["CMA_KPI_Label"]),
        ],
    ]
    kpi_table = Table(kpi_data, colWidths=[5.25 * cm] * 3)
    kpi_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DDE4EB")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E6EBF0")),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(kpi_table)
    story.append(Spacer(1, 0.45 * cm))

    comments = build_automatic_commentary(
        daylight_share,
        production_period_share,
        self_consumption_rate,
        self_sufficiency_rate,
    )
    story.append(Paragraph(comments["profile"], styles["CMA_Body"]))
    story.append(Paragraph(comments["self_consumption"], styles["CMA_Body"]))
    story.append(Paragraph(comments["self_sufficiency"], styles["CMA_Body"]))

    score_detail_rows = [
        ["Critère", "Poids", "Score"],
        ["Correspondance usages / production", "30 %", f"{cma_score_data.get('overlap_score', 0):.0f}/100"],
        ["Autoconsommation", "25 %", f"{cma_score_data.get('self_consumption_score', 0):.0f}/100"],
        ["Autoproduction", "20 %", f"{cma_score_data.get('self_sufficiency_score', 0):.0f}/100"],
        ["Régularité de la consommation", "15 %", f"{cma_score_data.get('regularity_score', 0):.0f}/100"],
        ["Potentiel solaire local", "10 %", f"{cma_score_data.get('solar_resource_score', 0):.0f}/100"],
    ]
    score_detail_table = Table(
        score_detail_rows,
        colWidths=[9.6 * cm, 2.5 * cm, 3.7 * cm],
    )
    score_detail_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), cma_blue),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(score_detail_table)
    story.append(
        Paragraph(
            "L'indice CMA est un outil pédagogique d'aide à la lecture. "
            "Il ne remplace ni l'étude technique, ni l'analyse économique.",
            styles["CMA_Small"],
        )
    )
    story.append(Spacer(1, 0.3 * cm))

    # Graphique mensuel
    fig_monthly_pdf = px.bar(
        monthly_df,
        x="Mois_date",
        y="Consommation_kWh",
        labels={"Mois_date": "Mois", "Consommation_kWh": "Consommation (kWh)"},
        title="Consommation mensuelle",
        color_discrete_sequence=["#17365D"],
    )
    fig_monthly_pdf.update_layout(template="plotly_white", showlegend=False)
    month_img = Image(figure_to_png_bytes(fig_monthly_pdf), width=16.5 * cm, height=7.4 * cm)
    story.append(month_img)
    story.append(PageBreak())

    # Profil hebdomadaire
    story.append(Paragraph("2. Comprendre le profil de consommation", styles["CMA_H1"]))
    story.append(
        Paragraph(
            "Le tableau ci-dessous présente la puissance moyenne appelée pour "
            "chaque heure et chaque jour de la semaine. Les valeurs les plus "
            "élevées signalent les périodes où l'activité consomme le plus.",
            styles["CMA_Body"],
        )
    )

    matrix = weekday_hour_matrix.copy().round(1)
    matrix_rows = [["Heure"] + list(matrix.columns)]
    for hour, row in matrix.iterrows():
        matrix_rows.append(
            [f"{int(hour):02d}h-{(int(hour)+1)%24:02d}h"]
            + ["" if pd.isna(v) else f"{v:.1f}" for v in row]
        )

    heat_table = Table(
        matrix_rows,
        colWidths=[2.25 * cm] + [1.85 * cm] * 7,
        repeatRows=1,
    )
    heat_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), cma_blue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.white),
        ("BACKGROUND", (0, 1), (0, -1), light_grey),
    ]

    numeric_values = matrix.to_numpy(dtype=float)
    valid_values = numeric_values[np.isfinite(numeric_values)]
    vmin = float(valid_values.min()) if valid_values.size else 0
    vmax = float(valid_values.max()) if valid_values.size else 1
    span = max(vmax - vmin, 1e-9)

    for r_index, (_, row) in enumerate(matrix.iterrows(), start=1):
        for c_index, value in enumerate(row, start=1):
            if pd.isna(value):
                bg = colors.white
            else:
                ratio = (float(value) - vmin) / span
                if ratio < 0.25:
                    bg = colors.HexColor("#63BE7B")
                elif ratio < 0.50:
                    bg = colors.HexColor("#A9D26D")
                elif ratio < 0.70:
                    bg = colors.HexColor("#FFEB84")
                elif ratio < 0.87:
                    bg = colors.HexColor("#F6B26B")
                else:
                    bg = colors.HexColor("#F8696B")
            heat_styles.append(("BACKGROUND", (c_index, r_index), (c_index, r_index), bg))

    heat_table.setStyle(TableStyle(heat_styles))
    story.append(heat_table)
    story.append(Spacer(1, 0.45 * cm))

    profile_long = (
        weekday_hour_matrix.reset_index()
        .melt(id_vars="Heure", var_name="Jour", value_name="Puissance_kW")
        .dropna()
    )
    fig_profiles_pdf = px.line(
        profile_long,
        x="Heure",
        y="Puissance_kW",
        color="Jour",
        title="Profil horaire moyen selon le jour de la semaine",
        labels={"Heure": "Heure", "Puissance_kW": "Puissance moyenne (kW)"},
    )
    fig_profiles_pdf.update_layout(template="plotly_white", hovermode="x unified")
    story.append(Image(figure_to_png_bytes(fig_profiles_pdf), width=16.5 * cm, height=7.6 * cm))
    story.append(PageBreak())

    # Solaire
    story.append(Paragraph("3. Potentiel solaire du site", styles["CMA_H1"]))
    solar_info = [
        ["Adresse analysée", safe_pdf_text(address_label or "Non renseignée")],
        ["Coordonnées", f"{latitude:.6f}, {longitude:.6f}" if latitude is not None else "Non disponibles"],
        ["Puissance étudiée", f"{pv_peak_kwp:g} kWc"],
        ["Orientation", safe_pdf_text(orientation_label)],
        ["Inclinaison", f"{pv_tilt:g}°"],
        ["Pertes prises en compte", f"{pv_losses:g} %"],
        ["Production PV estimée", f"{format_fr(pvgis_production_kwh, 0)} kWh" if not pd.isna(pvgis_production_kwh) else "Non disponible"],
        ["Productible estimé", f"{format_fr(annual_yield_kwh_per_kwp, 0)} kWh/kWc" if not pd.isna(annual_yield_kwh_per_kwp) else "Non disponible"],
        ["Énergie autoconsommée estimée", f"{format_fr(self_consumed_kwh, 0)} kWh" if not pd.isna(self_consumed_kwh) else "Non disponible"],
    ]
    solar_table = Table(solar_info, colWidths=[5.2 * cm, 10.6 * cm])
    solar_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), light_blue),
                ("TEXTCOLOR", (0, 0), (0, -1), cma_blue),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(solar_table)
    story.append(Spacer(1, 0.35 * cm))
    story.append(
        Paragraph(
            "PVGIS estime le potentiel solaire à partir de la localisation, "
            "de l'orientation, de l'inclinaison et de données climatiques de "
            "référence. Les résultats ne correspondent pas à une mesure réelle "
            "de chaque journée mais à une simulation adaptée au pré-diagnostic.",
            styles["CMA_Body"],
        )
    )

    if "Production_PV_kW" in filtered_df.columns:
        solar_plot_pdf = filtered_df[
            ["Horodate", "Puissance_kW", "Production_PV_kW"]
        ].copy()
        fig_compare_pdf = go.Figure()
        fig_compare_pdf.add_trace(
            go.Scatter(
                x=solar_plot_pdf["Horodate"],
                y=solar_plot_pdf["Puissance_kW"],
                name="Consommation",
                mode="lines",
                line=dict(color="#17365D", width=1.3),
            )
        )
        fig_compare_pdf.add_trace(
            go.Scatter(
                x=solar_plot_pdf["Horodate"],
                y=solar_plot_pdf["Production_PV_kW"],
                name="Production PV estimée",
                mode="lines",
                line=dict(color="#E53935", width=1.3),
            )
        )
        fig_compare_pdf.update_layout(
            title="Consommation et production photovoltaïque simulée",
            xaxis_title="Date",
            yaxis_title="Puissance (kW)",
            template="plotly_white",
            legend=dict(orientation="h"),
        )
        story.append(Image(figure_to_png_bytes(fig_compare_pdf), width=16.5 * cm, height=7.6 * cm))

    story.append(PageBreak())

    # Analyse tarifaire
    story.append(Paragraph("4. Analyse tarifaire HP / HC", styles["CMA_H1"]))
    story.append(
        Paragraph(
            "Cette analyse répartit les consommations selon les plages "
            "d'heures creuses renseignées et selon les saisons tarifaires : "
            "hiver du 1er novembre au 31 mars, été du 1er avril au 31 octobre.",
            styles["CMA_Body"],
        )
    )

    hc_labels = [
        f"{start.strftime('%H:%M')} - {end.strftime('%H:%M')}"
        for start, end in hc_ranges
    ]
    story.append(
        Paragraph(
            "<b>Plages d'heures creuses utilisées :</b> "
            + ", ".join(hc_labels),
            styles["CMA_Body"],
        )
    )

    tariff_pdf_rows = [["Catégorie", "Consommation", "Part"]]
    for _, row in tariff_summary_df.iterrows():
        tariff_pdf_rows.append(
            [
                str(row["Categorie_tarifaire"]),
                f"{format_fr(row['Consommation_kWh'], 0)} kWh",
                f"{format_fr(row['Part_pourcent'], 1)} %",
            ]
        )

    tariff_pdf_table = Table(
        tariff_pdf_rows,
        colWidths=[7.5 * cm, 4.2 * cm, 4.1 * cm],
    )
    tariff_pdf_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), cma_blue),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(tariff_pdf_table)
    story.append(Spacer(1, 0.35 * cm))

    story.append(
        Paragraph(
            f"<b>Indice de potentiel d'optimisation tarifaire CMA : "
            f"{tariff_score_data['score']:.0f}/100 — "
            f"{safe_pdf_text(tariff_score_data['label'])}</b>",
            styles["CMA_H2"],
        )
    )
    story.append(
        Paragraph(
            safe_pdf_text(build_tariff_commentary(tariff_score_data)),
            styles["CMA_Body"],
        )
    )
    story.append(
        Paragraph(
            "Cet indice signale un potentiel d'analyse. Il ne compare pas "
            "les prix des fournisseurs et ne prouve pas qu'un changement "
            "d'option tarifaire serait rentable.",
            styles["CMA_Small"],
        )
    )

    story.append(PageBreak())

    # Étude financière
    story.append(Paragraph("5. Étude financière indicative", styles["CMA_H1"]))
    story.append(
        Paragraph(
            safe_pdf_text(business_assistant["conclusion"]),
            styles["CMA_Body"],
        )
    )

    finance_kpis = [
        [
            Paragraph("Investissement net", styles["CMA_KPI_Label"]),
            Paragraph(
                f"{format_fr(investment_data['net_total'], 0)} € HT",
                styles["CMA_KPI_Value"],
            ),
            Paragraph("Gain net année 1", styles["CMA_KPI_Label"]),
            Paragraph(
                f"{format_fr(financial_projection['annual_net_gain_year_1'], 0)} €",
                styles["CMA_KPI_Value"],
            ),
        ],
        [
            Paragraph("Temps de retour", styles["CMA_KPI_Label"]),
            Paragraph(
                (
                    f"{format_fr(financial_projection['payback_year'], 1)} ans"
                    if not pd.isna(financial_projection["payback_year"])
                    else "Au-delà de l'horizon"
                ),
                styles["CMA_KPI_Value"],
            ),
            Paragraph("VAN", styles["CMA_KPI_Label"]),
            Paragraph(
                f"{format_fr(financial_projection['npv'], 0)} €",
                styles["CMA_KPI_Value"],
            ),
        ],
        [
            Paragraph("TRI indicatif", styles["CMA_KPI_Label"]),
            Paragraph(
                (
                    f"{format_fr(financial_projection['irr'] * 100, 1)} %"
                    if not pd.isna(financial_projection["irr"])
                    else "Non calculable"
                ),
                styles["CMA_KPI_Value"],
            ),
            Paragraph(
                f"Gain cumulé à {financial_horizon_years} ans",
                styles["CMA_KPI_Label"],
            ),
            Paragraph(
                f"{format_fr(financial_projection['total_net_gain'], 0)} €",
                styles["CMA_KPI_Value"],
            ),
        ],
    ]
    finance_kpi_table = Table(
        finance_kpis,
        colWidths=[3.6 * cm, 4.3 * cm, 3.6 * cm, 4.3 * cm],
    )
    finance_kpi_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), light_grey),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D8E0E8")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(finance_kpi_table)
    story.append(Spacer(1, 0.35 * cm))

    story.append(Paragraph("Répartition des coûts", styles["CMA_H2"]))
    cost_rows = [["Poste", "Montant", "Part"]]
    cost_items = [
        ("Modules, onduleur et pose", investment_data["equipment_cost"]),
        ("Système de fixation", investment_data["fixing_cost"]),
        ("Surcoût ERP / ICPE", investment_data["erp_surcharge_cost"]),
        ("Étude structure", investment_data["structural_study_cost"]),
        ("Rénovation de couverture", investment_data["roof_cost"]),
        ("Désamiantage", investment_data["asbestos_cost"]),
        ("Raccordement", investment_data["connection_cost"]),
        ("Autres coûts", investment_data["other_investment_costs"]),
    ]
    positive_costs = [
        (label, amount)
        for label, amount in cost_items
        if amount > 0
    ]
    gross_cost = max(investment_data["gross_total"], 1)
    for label, amount in positive_costs:
        cost_rows.append(
            [
                label,
                f"{format_fr(amount, 0)} €",
                f"{format_fr(amount / gross_cost * 100, 1)} %",
            ]
        )
    cost_rows.extend(
        [
            ["Investissement brut", f"{format_fr(investment_data['gross_total'], 0)} €", "100 %"],
            ["Aides déduites", f"- {format_fr(investment_data['grant_amount'], 0)} €", ""],
            ["Investissement net", f"{format_fr(investment_data['net_total'], 0)} €", ""],
        ]
    )
    cost_table = Table(
        cost_rows,
        colWidths=[9.0 * cm, 4.0 * cm, 2.8 * cm],
    )
    cost_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), cma_blue),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(cost_table)

    story.append(PageBreak())
    story.append(
        Paragraph(
            f"6. Prévisionnel financier sur {financial_horizon_years} ans",
            styles["CMA_H1"],
        )
    )

    projection_table_df = financial_projection["table"].copy()
    selected_years = sorted(
        set(
            [
                1,
                min(5, financial_horizon_years),
                min(10, financial_horizon_years),
                min(15, financial_horizon_years),
                financial_horizon_years,
            ]
        )
    )
    forecast_rows = [
        ["Année", "Flux net", "Cumul net", "Cumul actualisé"]
    ]
    for year in selected_years:
        row = projection_table_df[
            projection_table_df["Année"] == year
        ].iloc[0]
        forecast_rows.append(
            [
                str(year),
                f"{format_fr(row['Flux net (€)'], 0)} €",
                f"{format_fr(row['Cumul net (€)'], 0)} €",
                f"{format_fr(row['Cumul actualisé (€)'], 0)} €",
            ]
        )
    forecast_table = Table(
        forecast_rows,
        colWidths=[2.4 * cm, 4.4 * cm, 4.5 * cm, 4.5 * cm],
    )
    forecast_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), cma_blue),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(forecast_table)
    story.append(Spacer(1, 0.35 * cm))

    assumptions_rows = [
        ["Hypothèse", "Valeur"],
        ["Tarification électrique", safe_pdf_text(electricity_tariff_type)],
        ["Vente du surplus", f"{format_fr(surplus_sale_price_eur_kwh, 4)} €/kWh HT"],
        ["Hausse du prix de l'électricité", f"{format_fr(electricity_price_increase_percent, 1)} %/an"],
        ["Dégradation de production", f"{format_fr(production_degradation_percent, 1)} %/an"],
        ["Taux d'actualisation", f"{format_fr(discount_rate_percent, 1)} %"],
        ["Charges annuelles initiales", f"{format_fr(operating_cost_data['total'], 0)} € HT/an"],
    ]
    assumptions_table = Table(
        assumptions_rows,
        colWidths=[9.2 * cm, 6.6 * cm],
    )
    assumptions_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), cma_blue),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(Paragraph("Hypothèses financières", styles["CMA_H2"]))
    story.append(assumptions_table)

    story.append(PageBreak())
    story.append(Paragraph("7. Synthèse de l'assistant CMA", styles["CMA_H1"]))
    story.append(
        Paragraph(
            f"<b>{safe_pdf_text(business_assistant['headline'])}</b>",
            styles["CMA_H2"],
        )
    )
    story.append(
        Paragraph(
            safe_pdf_text(business_assistant["conclusion"]),
            styles["CMA_Body"],
        )
    )
    story.append(Paragraph("Points favorables", styles["CMA_H2"]))
    for item in business_assistant["strengths"]:
        story.append(
            Paragraph("• " + safe_pdf_text(item), styles["CMA_Body"])
        )
    story.append(Paragraph("Points de vigilance", styles["CMA_H2"]))
    for item in business_assistant["vigilance"]:
        story.append(
            Paragraph("• " + safe_pdf_text(item), styles["CMA_Body"])
        )
    story.append(Paragraph("Prochaines étapes", styles["CMA_H2"]))
    for number, item in enumerate(
        business_assistant["next_steps"],
        start=1,
    ):
        story.append(
            Paragraph(
                f"{number}. {safe_pdf_text(item)}",
                styles["CMA_Body"],
            )
        )

    story.append(PageBreak())

    # Explications
    story.append(Paragraph("8. Comment lire les résultats ?", styles["CMA_H1"]))
    explanations = [
        (
            "Autoconsommation",
            "Part de l'électricité solaire produite qui serait consommée "
            "directement par l'entreprise au moment où elle est produite.",
        ),
        (
            "Autoproduction",
            "Part des besoins électriques de l'entreprise qui pourrait être "
            "couverte par les panneaux photovoltaïques.",
        ),
        (
            "Surplus",
            "Électricité produite mais non consommée immédiatement. Elle peut, "
            "selon le projet, être injectée sur le réseau ou faire l'objet "
            "d'une autre stratégie de valorisation.",
        ),
        (
            "Électricité achetée au réseau",
            "Part de la consommation qui reste fournie par le réseau, notamment "
            "la nuit ou lorsque la production solaire est insuffisante.",
        ),
    ]
    for title, text in explanations:
        story.append(Paragraph(title, styles["CMA_H2"]))
        story.append(Paragraph(text, styles["CMA_Body"]))

    story.append(Paragraph("9. Recommandations générales", styles["CMA_H1"]))
    recommendations = [
        "Vérifier la surface réellement disponible et les zones d'ombrage.",
        "Faire contrôler l'état et la capacité portante de la toiture.",
        "Comparer plusieurs puissances d'installation avant de retenir un scénario.",
        "Intégrer les coûts, aides, tarifs d'achat et conditions de raccordement.",
        "Faire confirmer le projet par un installateur qualifié et un bureau d'étude si nécessaire.",
    ]
    recommendation_rows = [
        [Paragraph("• " + item, styles["CMA_Body"])]
        for item in recommendations
    ]
    story.append(Table(recommendation_rows, colWidths=[15.8 * cm]))

    story.append(Paragraph("10. Limites du pré-diagnostic", styles["CMA_H1"]))
    story.append(
        Paragraph(
            "Ce rapport repose sur les données Enedis importées et sur une "
            "simulation PVGIS. Il ne vérifie pas la structure du bâtiment, "
            "les contraintes d'urbanisme, le raccordement, les ombrages fins, "
            "les conditions fiscales ou le financement bancaire. "
            "Les résultats doivent être confirmés dans le cadre d'une étude complète.",
            styles["CMA_Body"],
        )
    )
    story.append(Spacer(1, 0.4 * cm))
    story.append(
        Paragraph(
            f"Fichier analysé : {safe_pdf_text(source_filename)} - "
            f"unité source : {safe_pdf_text(source_unit)} - "
            f"pas source : {safe_pdf_text(time_step)}.",
            styles["CMA_Small"],
        )
    )

    doc.build(story)
    output.seek(0)
    return output.getvalue()


def format_fr(value: float, decimals: int = 1) -> str:
    return (
        f"{value:,.{decimals}f}"
        .replace(",", " ")
        .replace(".", ",")
    )


def make_excel_export(
    hourly_standardized_data: pd.DataFrame,
    hourly_data: pd.DataFrame,
    daily_data: pd.DataFrame,
    monthly_data: pd.DataFrame,
    weekday_hour_matrix: pd.DataFrame,
    tariff_summary: pd.DataFrame,
    tariff_detail: pd.DataFrame,
    financial_summary: pd.DataFrame,
    financial_projection_data: pd.DataFrame,
    summary: pd.DataFrame,
) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Synthèse")
        hourly_standardized_data.to_excel(
            writer,
            index=False,
            sheet_name="Données traitées 1h",
        )
        hourly_data.to_excel(writer, index=False, sheet_name="Profil horaire")
        daily_data.to_excel(
            writer,
            index=False,
            sheet_name="Consommations journalières",
        )
        monthly_data.to_excel(
            writer,
            index=False,
            sheet_name="Consommations mensuelles",
        )
        weekday_hour_matrix.to_excel(
            writer,
            sheet_name="Moyenne heure-jour",
        )
        tariff_summary.to_excel(
            writer,
            index=False,
            sheet_name="Répartition tarifaire",
        )
        tariff_detail.to_excel(
            writer,
            index=False,
            sheet_name="Détail tarifaire",
        )
        financial_summary.to_excel(
            writer,
            index=False,
            sheet_name="Synthèse financière",
        )
        financial_projection_data.to_excel(
            writer,
            index=False,
            sheet_name="Projection financière",
        )

        workbook = writer.book

        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    max_length = max(
                        max_length,
                        len(str(cell.value or "")),
                    )

                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 11),
                    35,
                )

    return output.getvalue()


# ============================================================
# ACCUEIL
# ============================================================

render_header()

st.markdown(
    """
    <div class="intro-card">
        <div class="intro-icon">⚡</div>
        <div>
            <strong>Analysez un fichier Enedis en quelques secondes.</strong><br>
            L'application transforme automatiquement les courbes de charge en
            profils horaires, consommations journalières, tableaux thermiques
            et indicateurs utiles au pré-diagnostic photovoltaïque.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# IMPORT ET PARAMÈTRES
# ============================================================

with st.sidebar:
    st.markdown("## 1. Import")

    uploaded_file = st.file_uploader(
        "Fichier Enedis",
        type=["csv", "xlsx", "xls"],
        help="Colonnes obligatoires : Horodate et Valeur.",
    )

if uploaded_file is None:
    st.markdown(
        """
        <div class="feature-grid">
            <div class="feature-card">
                <div class="feature-icon">📂</div>
                <h3>1. Importer</h3>
                <p>
                    Déposez un export Enedis au format CSV ou Excel
                    depuis le panneau latéral.
                </p>
            </div>
            <div class="feature-card">
                <div class="feature-icon">📊</div>
                <h3>2. Analyser</h3>
                <p>
                    Obtenez les profils de charge, consommations,
                    tableaux thermiques et contrôles qualité.
                </p>
            </div>
            <div class="feature-card">
                <div class="feature-icon">📥</div>
                <h3>3. Exporter</h3>
                <p>
                    Téléchargez un classeur Excel structuré,
                    prêt à être exploité dans le diagnostic.
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.stop()


try:
    source_df = read_enedis_file(
        uploaded_file.getvalue(),
        uploaded_file.name,
    )
    time_step = detect_time_step(source_df)
    source_unit = detect_source_unit(source_df)
    enriched_df, interpretation = enrich_energy_data(
        source_df,
        time_step,
        source_unit,
    )
except Exception as exc:
    st.error(f"Impossible de traiter le fichier : {exc}")
    st.stop()


available_years = sorted(
    enriched_df["Horodate"].dt.year.unique().tolist()
)

with st.sidebar:
    st.markdown("---")
    st.markdown("## 2. Période")

    period_mode = st.radio(
        "Période analysée",
        [
            "Toutes les données",
            "Année",
            "Période personnalisée",
        ],
    )

    selected_year = None
    start_date = enriched_df["Horodate"].min().date()
    end_date = enriched_df["Horodate"].max().date()

    if period_mode == "Année":
        selected_year = st.selectbox(
            "Année",
            available_years,
            index=len(available_years) - 1,
        )

    elif period_mode == "Période personnalisée":
        start_date = st.date_input(
            "Date de début",
            value=start_date,
        )
        end_date = st.date_input(
            "Date de fin",
            value=end_date,
        )

    st.markdown("---")
    st.markdown("## 3. Entreprise")

    company_name = st.text_input(
        "Nom de l'entreprise",
        placeholder="Ex. Atelier Dupont",
    )

    company_siret = st.text_input(
        "SIRET (optionnel)",
        placeholder="Ex. 123 456 789 00012",
    )

    advisor_name = st.text_input(
        "Conseiller CMA",
        placeholder="Nom du conseiller",
    )

    diagnostic_date = st.date_input(
        "Date du diagnostic",
        value=pd.Timestamp.today().date(),
    )

    st.markdown("---")
    st.markdown("## 4. Localisation solaire")

    company_address = st.text_input(
        "Adresse complète de l'entreprise",
        placeholder="Ex. 3 rue du 11 Novembre, 33000 Bordeaux",
    )

    if st.button(
        "📍 Rechercher l'adresse",
        use_container_width=True,
    ):
        try:
            st.session_state["address_candidates"] = (
                geocode_company_address(company_address)
            )
        except Exception as exc:
            st.session_state["address_candidates"] = []
            st.error(f"Géocodage impossible : {exc}")

    address_candidates = st.session_state.get(
        "address_candidates",
        [],
    )

    selected_location = None

    if address_candidates:
        labels = [item["label"] for item in address_candidates]

        selected_label = st.selectbox(
            "Adresse reconnue",
            labels,
        )

        selected_location = next(
            item
            for item in address_candidates
            if item["label"] == selected_label
        )

        st.success(
            f"Adresse validée : {selected_location['label']}\n\n"
            f"Latitude : {selected_location['latitude']:.6f}\n\n"
            f"Longitude : {selected_location['longitude']:.6f}"
        )

    with st.expander("Coordonnées manuelles en cas de besoin"):
        manual_coordinates = st.checkbox(
            "Utiliser des coordonnées manuelles"
        )

        manual_latitude = st.number_input(
            "Latitude",
            min_value=-90.0,
            max_value=90.0,
            value=44.8378,
            format="%.6f",
            disabled=not manual_coordinates,
        )

        manual_longitude = st.number_input(
            "Longitude",
            min_value=-180.0,
            max_value=180.0,
            value=-0.5792,
            format="%.6f",
            disabled=not manual_coordinates,
        )

        if manual_coordinates:
            selected_location = {
                "label": (
                    company_address.strip()
                    or "Coordonnées saisies manuellement"
                ),
                "latitude": float(manual_latitude),
                "longitude": float(manual_longitude),
                "score": None,
                "postcode": "",
                "city": "",
                "street": "",
                "housenumber": "",
                "source": "Saisie manuelle",
            }

    st.markdown("## 5. Paramètres photovoltaïques")

    pv_peak_kwp = st.number_input(
        "Puissance étudiée (kWc)",
        min_value=0.1,
        max_value=1000.0,
        value=10.0,
        step=0.5,
    )

    pv_tilt = st.slider(
        "Inclinaison des panneaux",
        min_value=0,
        max_value=90,
        value=30,
        format="%d°",
    )

    orientation_label = st.selectbox(
        "Orientation",
        [
            "Sud",
            "Sud-Est",
            "Est",
            "Sud-Ouest",
            "Ouest",
        ],
    )

    orientation_to_pvgis = {
        "Sud": 0,
        "Sud-Est": -45,
        "Est": -90,
        "Sud-Ouest": 45,
        "Ouest": 90,
    }

    pv_aspect = orientation_to_pvgis[orientation_label]

    pv_losses = st.slider(
        "Pertes système",
        min_value=0,
        max_value=35,
        value=14,
        format="%d%%",
    )

    st.caption(
        "Le lever et le coucher sont calculés pour chaque date à partir "
        "des coordonnées exactes. PVGIS estime automatiquement "
        "l'irradiation et la production pour la localisation et les "
        "caractéristiques de l'installation."
    )

    st.markdown("---")
    st.markdown("## 6. Paramètres tarifaires")

    hc_range_count = st.radio(
        "Nombre de plages d'heures creuses par jour",
        [1, 2],
        horizontal=True,
        help=(
            "Renseignez uniquement les plages d'heures creuses. "
            "Toutes les autres heures seront automatiquement classées "
            "en heures pleines."
        ),
    )

    hc1_col1, hc1_col2 = st.columns(2)

    with hc1_col1:
        hc1_start = st.time_input(
            "Début HC 1",
            value=pd.Timestamp("22:00").time(),
            step=1800,
        )

    with hc1_col2:
        hc1_end = st.time_input(
            "Fin HC 1",
            value=pd.Timestamp("06:00").time(),
            step=1800,
        )

    hc_ranges = [(hc1_start, hc1_end)]

    if hc_range_count == 2:
        hc2_col1, hc2_col2 = st.columns(2)

        with hc2_col1:
            hc2_start = st.time_input(
                "Début HC 2",
                value=pd.Timestamp("12:30").time(),
                step=1800,
            )

        with hc2_col2:
            hc2_end = st.time_input(
                "Fin HC 2",
                value=pd.Timestamp("14:30").time(),
                step=1800,
            )

        hc_ranges.append((hc2_start, hc2_end))

    st.caption(
        "Hiver / saison haute : du 1er novembre au 31 mars. "
        "Été / saison basse : du 1er avril au 31 octobre. "
        "Le milieu de chaque intervalle Enedis est utilisé pour le classement."
    )

    st.markdown("---")
    st.markdown("## 7. Hypothèses économiques")

    electricity_tariff_type = st.selectbox(
        "Type de tarif d'achat",
        [
            "Tarif unique",
            "HP / HC",
            "HP / HC hiver-été",
        ],
        help=(
            "Saisissez les prix présents sur une facture ou un contrat récent. "
            "Les tarifs ne sont pas maintenus automatiquement par l'application."
        ),
    )

    unique_electricity_price = 0.18
    hp_electricity_price = 0.20
    hc_electricity_price = 0.15
    hp_winter_electricity_price = 0.21
    hc_winter_electricity_price = 0.16
    hp_summer_electricity_price = 0.18
    hc_summer_electricity_price = 0.14

    if electricity_tariff_type == "Tarif unique":
        unique_electricity_price = st.number_input(
            "Prix d'achat de l'électricité (€/kWh HT)",
            min_value=0.0,
            max_value=5.0,
            value=0.1842,
            step=0.0010,
            format="%.4f",
        )

    elif electricity_tariff_type == "HP / HC":
        hp_col, hc_col = st.columns(2)

        with hp_col:
            hp_electricity_price = st.number_input(
                "Prix HP (€/kWh HT)",
                min_value=0.0,
                max_value=5.0,
                value=0.20,
                step=0.0010,
                format="%.4f",
            )

        with hc_col:
            hc_electricity_price = st.number_input(
                "Prix HC (€/kWh HT)",
                min_value=0.0,
                max_value=5.0,
                value=0.15,
                step=0.0010,
                format="%.4f",
            )

    else:
        hp_winter_electricity_price = st.number_input(
            "Prix HP hiver (€/kWh HT)",
            min_value=0.0,
            max_value=5.0,
            value=0.21,
            step=0.0010,
            format="%.4f",
        )
        hc_winter_electricity_price = st.number_input(
            "Prix HC hiver (€/kWh HT)",
            min_value=0.0,
            max_value=5.0,
            value=0.16,
            step=0.0010,
            format="%.4f",
        )
        hp_summer_electricity_price = st.number_input(
            "Prix HP été (€/kWh HT)",
            min_value=0.0,
            max_value=5.0,
            value=0.18,
            step=0.0010,
            format="%.4f",
        )
        hc_summer_electricity_price = st.number_input(
            "Prix HC été (€/kWh HT)",
            min_value=0.0,
            max_value=5.0,
            value=0.14,
            step=0.0010,
            format="%.4f",
        )

    annual_subscription_eur = st.number_input(
        "Abonnement annuel (€ HT)",
        min_value=0.0,
        max_value=100000.0,
        value=300.0,
        step=10.0,
    )

    surplus_sale_price_eur_kwh = st.number_input(
        "Prix de vente du surplus (€/kWh HT)",
        min_value=0.0,
        max_value=5.0,
        value=0.0761,
        step=0.0010,
        format="%.4f",
    )

    st.caption(
        "Les paramètres détaillés d'investissement, de raccordement et de "
        "projection sont regroupés dans le bouton ci-dessous."
    )

    with st.popover(
        "⚙️ Paramètres financiers avancés",
        use_container_width=True,
    ):
        st.markdown("### Investissement")

        fixing_type = st.selectbox(
            "Système de fixation",
            list(PV_FIXING_COSTS_EUR_WC.keys()),
        )

        erp_icpe_surcharge = st.checkbox(
            "Projet ERP ou ICPE : ajouter 0,10 €/Wc",
            value=False,
        )

        structural_study_cost = st.number_input(
            "Étude structure charpente/toiture (€ HT)",
            min_value=0.0,
            max_value=100000.0,
            value=2000.0,
            step=250.0,
        )

        roof_renovation_enabled = st.checkbox(
            "Intégrer une rénovation de couverture",
            value=False,
        )

        roof_type = st.selectbox(
            "Type de couverture à rénover",
            list(ROOF_RENOVATION_COSTS_EUR_M2.keys()),
            disabled=not roof_renovation_enabled,
        )

        roof_area_m2 = st.number_input(
            "Surface de couverture concernée (m²)",
            min_value=0.0,
            max_value=100000.0,
            value=0.0,
            step=10.0,
            disabled=not roof_renovation_enabled,
        )

        asbestos_removal_enabled = st.checkbox(
            "Prévoir un désamiantage à 60 €/m²",
            value=False,
            disabled=not roof_renovation_enabled,
        )

        st.markdown("### Raccordement")

        connection_mode = st.selectbox(
            "Scénario indicatif de raccordement",
            [
                "Aucun / inférieur ou égal à 36 kWc",
                "Branchement BT avec extension",
                "Branchement complet C4",
                "Création poste BT-HTA",
            ],
        )

        public_extension_length_m = st.number_input(
            "Longueur d'extension publique (m)",
            min_value=0.0,
            max_value=10000.0,
            value=0.0,
            step=5.0,
            disabled=(
                connection_mode
                not in [
                    "Branchement BT avec extension",
                    "Création poste BT-HTA",
                ]
            ),
        )

        apply_enedis_reduction = st.checkbox(
            "Appliquer une réfaction Enedis indicative de 60 %",
            value=True,
            disabled=pv_peak_kwp <= 36,
            help=(
                "Hypothèse indicative issue du document transmis. "
                "Elle doit être confirmée par une proposition de raccordement."
            ),
        )

        private_trench_length_m = st.number_input(
            "Tranchée interne privée (m)",
            min_value=0.0,
            max_value=10000.0,
            value=0.0,
            step=5.0,
        )

        include_private_hta_post = st.checkbox(
            "Création d'un poste privé HTA/BT",
            value=False,
        )

        include_decoupling_cell = st.checkbox(
            "Ajouter une cellule de découplage",
            value=False,
        )

        other_investment_costs = st.number_input(
            "Autres coûts d'investissement (€ HT)",
            min_value=0.0,
            max_value=10000000.0,
            value=0.0,
            step=500.0,
        )

        grant_amount = st.number_input(
            "Aides ou subventions déduites (€)",
            min_value=0.0,
            max_value=10000000.0,
            value=0.0,
            step=500.0,
        )

        st.markdown("### Charges annuelles")

        insurance_rate_percent = st.slider(
            "Assurance multirisque et perte d'exploitation",
            min_value=0.0,
            max_value=2.0,
            value=0.5,
            step=0.1,
            format="%.1f %% de l'investissement",
        )

        maintenance_eur_kwp = st.slider(
            "Suivi et maintenance",
            min_value=0.0,
            max_value=30.0,
            value=10.5,
            step=0.5,
            format="%.1f €/kWc/an",
        )

        inverter_provision_eur_kwp = st.slider(
            "Provision remplacement onduleurs",
            min_value=0.0,
            max_value=15.0,
            value=3.0,
            step=0.5,
            format="%.1f €/kWc/an",
        )

        ifer_rate_eur_kwp = st.number_input(
            "Tarif IFER si puissance > 100 kWc (€/kWc/an)",
            min_value=0.0,
            max_value=100.0,
            value=3.542,
            step=0.001,
            format="%.3f",
        )

        other_annual_costs = st.number_input(
            "Autres charges annuelles (€ HT/an)",
            min_value=0.0,
            max_value=1000000.0,
            value=0.0,
            step=100.0,
        )

        st.markdown("### Projection")

        financial_horizon_years = st.slider(
            "Durée de projection",
            min_value=5,
            max_value=40,
            value=20,
            step=1,
            format="%d ans",
        )

        electricity_price_increase_percent = st.slider(
            "Hausse annuelle du prix d'achat",
            min_value=0.0,
            max_value=10.0,
            value=2.0,
            step=0.1,
            format="%.1f %%",
        )

        surplus_price_increase_percent = st.slider(
            "Évolution annuelle du tarif de surplus",
            min_value=-5.0,
            max_value=10.0,
            value=0.0,
            step=0.1,
            format="%.1f %%",
        )

        production_degradation_percent = st.slider(
            "Dégradation annuelle de la production",
            min_value=0.0,
            max_value=2.0,
            value=0.5,
            step=0.1,
            format="%.1f %%",
        )

        operating_cost_increase_percent = st.slider(
            "Hausse annuelle des charges",
            min_value=0.0,
            max_value=10.0,
            value=2.0,
            step=0.1,
            format="%.1f %%",
        )

        discount_rate_percent = st.slider(
            "Taux d'actualisation",
            min_value=0.0,
            max_value=15.0,
            value=4.0,
            step=0.1,
            format="%.1f %%",
        )

        st.caption(
            "Les coûts sont des ordres de grandeur HT issus des documents "
            "métier transmis. Ils doivent être confirmés par des devis, une "
            "étude structure et une proposition de raccordement Enedis."
        )



filtered_df = filter_period(
    enriched_df,
    period_mode,
    selected_year,
    start_date,
    end_date,
)

if filtered_df.empty:
    st.warning("Aucune donnée sur la période sélectionnée.")
    st.stop()


# ============================================================
# CALCULS
# ============================================================

hourly_df = build_hourly_data(filtered_df)
daily_df = build_daily_data(filtered_df)
monthly_df = build_monthly_data(filtered_df)
weekday_hour_matrix = build_weekday_hour_matrix(hourly_df)

filtered_df = add_tariff_categories(
    filtered_df,
    hc_ranges,
)
tariff_summary_df = build_tariff_summary(filtered_df)

analysis_start = filtered_df["Horodate"].min()
analysis_end = filtered_df["Horodate"].max()
analysis_days = max(
    (analysis_end - analysis_start).total_seconds() / 86400,
    1,
)
coverage_ratio = min(analysis_days / 365.25, 1.0)

tariff_score_data = calculate_tariff_optimization_score(
    tariff_summary=tariff_summary_df,
    daily_consumption=daily_df["Consommation_kWh"],
    coverage_ratio=coverage_ratio,
)

tariff_values = tariff_summary_df.set_index(
    "Categorie_tarifaire"
)["Consommation_kWh"]

hp_winter_kwh = float(tariff_values.get("HP hiver", 0))
hc_winter_kwh = float(tariff_values.get("HC hiver", 0))
hp_summer_kwh = float(tariff_values.get("HP été", 0))
hc_summer_kwh = float(tariff_values.get("HC été", 0))

total_kwh = filtered_df["Energie_kWh"].sum()
average_daily_kwh = daily_df["Consommation_kWh"].mean()
median_daily_kwh = daily_df["Consommation_kWh"].median()
maximum_power_kw = filtered_df["Puissance_kW"].max()
mean_power_kw = filtered_df["Puissance_kW"].mean()

solar_analysis_available = selected_location is not None
pvgis_available = False
solar_error = None
solar_daily_df = pd.DataFrame()

daylight_kwh = np.nan
daylight_share = np.nan
solar_rows_count = 0
solar_day_rows_count = 0
solar_event_rows_count = 0
solar_coherence_rate = np.nan
production_period_kwh = np.nan
production_period_share = np.nan
pvgis_production_kwh = np.nan
self_consumed_kwh = np.nan
pv_surplus_kwh = np.nan
grid_import_kwh = np.nan
self_consumption_rate = np.nan
self_sufficiency_rate = np.nan
annual_yield_kwh_per_kwp = np.nan
cma_score_data = {
    "score": 0.0,
    "label": "Non calculé",
    "color": "#7B8794",
    "overlap_score": 0.0,
    "self_consumption_score": 0.0,
    "self_sufficiency_score": 0.0,
    "regularity_score": 0.0,
    "solar_resource_score": 0.0,
    "coefficient_variation": np.nan,
}

if solar_analysis_available:
    try:
        filtered_df = add_astronomical_solar_data(
            filtered_df,
            latitude=selected_location["latitude"],
            longitude=selected_location["longitude"],
        )

        daylight_kwh = filtered_df.loc[
            filtered_df["Soleil_leve"],
            "Energie_kWh",
        ].sum()

        daylight_share = (
            daylight_kwh / total_kwh * 100
            if total_kwh
            else 0
        )

        solar_rows_count = len(filtered_df)
        solar_day_rows_count = int(
            filtered_df["Soleil_leve"].sum()
        )
        solar_event_rows_count = int(
            filtered_df["Lever_soleil"].notna().sum()
        )
        solar_coherence_rate = (
            filtered_df["Controle_solaire_coherent"].mean() * 100
            if solar_rows_count
            else np.nan
        )

        try:
            pvgis_profile, pvgis_metadata = (
                fetch_pvgis_reference_profile(
                    latitude=selected_location["latitude"],
                    longitude=selected_location["longitude"],
                    tilt=pv_tilt,
                    aspect=pv_aspect,
                    peak_power_kwp=pv_peak_kwp,
                    losses_percent=pv_losses,
                )
            )

            filtered_df = merge_pvgis_profile(
                filtered_df,
                pvgis_profile,
            )
            pvgis_available = True

            production_active_mask = (
                filtered_df["Production_PV_kW"].fillna(0) > 0
            )

            production_period_kwh = filtered_df.loc[
                production_active_mask,
                "Energie_kWh",
            ].sum()

            production_period_share = (
                production_period_kwh / total_kwh * 100
                if total_kwh
                else 0
            )

            pvgis_production_kwh = filtered_df[
                "Production_PV_kWh"
            ].sum()

            self_consumed_kwh = filtered_df[
                "Autoconsommation_estimee_kWh"
            ].sum()

            self_consumption_rate = (
                self_consumed_kwh / pvgis_production_kwh * 100
                if pvgis_production_kwh
                else 0
            )

            self_sufficiency_rate = (
                self_consumed_kwh / total_kwh * 100
                if total_kwh
                else 0
            )

            pv_surplus_kwh = max(
                pvgis_production_kwh - self_consumed_kwh,
                0,
            )

            grid_import_kwh = max(
                total_kwh - self_consumed_kwh,
                0,
            )

            annual_yield_kwh_per_kwp = (
                pvgis_production_kwh / pv_peak_kwp
                if pv_peak_kwp
                else np.nan
            )

            cma_score_data = calculate_cma_pv_score(
                production_period_share=production_period_share,
                self_consumption_rate=self_consumption_rate,
                self_sufficiency_rate=self_sufficiency_rate,
                daily_consumption=daily_df["Consommation_kWh"],
                annual_yield_kwh_per_kwp=annual_yield_kwh_per_kwp,
            )

            solar_daily_df = build_daily_solar_summary(
                filtered_df
            )

        except Exception as exc:
            solar_error = (
                "Les heures de lever/coucher ont été calculées, "
                f"mais PVGIS n'a pas pu être interrogé : {exc}"
            )

    except Exception as exc:
        solar_analysis_available = False
        solar_error = f"Analyse solaire impossible : {exc}"

analysis_duration_years = max(
    analysis_days / 365.25,
    1 / 365.25,
)

connection_data = calculate_connection_cost(
    peak_power_kwp=pv_peak_kwp,
    connection_mode=connection_mode,
    public_extension_length_m=public_extension_length_m,
    private_trench_length_m=private_trench_length_m,
    apply_enedis_reduction=apply_enedis_reduction,
    include_private_hta_post=include_private_hta_post,
    include_decoupling_cell=include_decoupling_cell,
)

investment_data = calculate_investment_costs(
    peak_power_kwp=pv_peak_kwp,
    fixing_type=fixing_type,
    erp_icpe_surcharge=erp_icpe_surcharge,
    structural_study_cost=structural_study_cost,
    roof_renovation_enabled=roof_renovation_enabled,
    roof_type=roof_type,
    roof_area_m2=roof_area_m2,
    asbestos_removal_enabled=asbestos_removal_enabled,
    connection_data=connection_data,
    other_investment_costs=other_investment_costs,
    grant_amount=grant_amount,
)

electricity_prices = tariff_price_map(
    tariff_type=electricity_tariff_type,
    unique_price=unique_electricity_price,
    hp_price=hp_electricity_price,
    hc_price=hc_electricity_price,
    hp_winter_price=hp_winter_electricity_price,
    hc_winter_price=hc_winter_electricity_price,
    hp_summer_price=hp_summer_electricity_price,
    hc_summer_price=hc_summer_electricity_price,
)

energy_value_data = calculate_energy_value(
    df=filtered_df,
    price_map=electricity_prices,
    surplus_sale_price=surplus_sale_price_eur_kwh,
    annual_subscription=annual_subscription_eur,
    analysis_years=analysis_duration_years,
)

operating_cost_data = calculate_annual_operating_costs(
    peak_power_kwp=pv_peak_kwp,
    investment_gross=investment_data["gross_total"],
    insurance_rate_percent=insurance_rate_percent,
    maintenance_eur_kwp=maintenance_eur_kwp,
    inverter_provision_eur_kwp=inverter_provision_eur_kwp,
    ifer_rate_eur_kwp=ifer_rate_eur_kwp,
    other_annual_costs=other_annual_costs,
)

financial_projection = build_financial_projection(
    net_investment=investment_data["net_total"],
    annual_self_consumption_saving=(
        energy_value_data["annual_self_consumption_saving"]
    ),
    annual_surplus_revenue=(
        energy_value_data["annual_surplus_revenue"]
    ),
    annual_operating_cost=operating_cost_data["total"],
    horizon_years=financial_horizon_years,
    electricity_price_increase_percent=(
        electricity_price_increase_percent
    ),
    surplus_price_increase_percent=(
        surplus_price_increase_percent
    ),
    production_degradation_percent=(
        production_degradation_percent
    ),
    operating_cost_increase_percent=(
        operating_cost_increase_percent
    ),
    discount_rate_percent=discount_rate_percent,
)

business_assistant = build_cma_business_assistant(
    daylight_share=daylight_share,
    production_period_share=production_period_share,
    self_consumption_rate=self_consumption_rate,
    self_sufficiency_rate=self_sufficiency_rate,
    pv_surplus_kwh=pv_surplus_kwh,
    pvgis_production_kwh=pvgis_production_kwh,
    cma_score_data=cma_score_data,
    tariff_score_data=tariff_score_data,
    investment_data=investment_data,
    operating_cost_data=operating_cost_data,
    financial_projection=financial_projection,
    roof_renovation_enabled=roof_renovation_enabled,
    asbestos_removal_enabled=asbestos_removal_enabled,
    connection_data=connection_data,
    coverage_ratio=coverage_ratio,
)

load_factor = (
    mean_power_kw / maximum_power_kw * 100
    if maximum_power_kw
    else 0
)

duplicate_count = int(
    filtered_df["Horodate"].duplicated().sum()
)

expected_points_per_day = round(
    pd.Timedelta(days=1) / time_step
)

one_hour_points = round(
    pd.Timedelta(hours=1) / time_step
)

points_per_day = (
    filtered_df.assign(
        Date=filtered_df["Horodate"].dt.date
    )
    .groupby("Date")
    .size()
)

valid_daily_counts = {
    expected_points_per_day,
    expected_points_per_day - one_hour_points,
    expected_points_per_day + one_hour_points,
}

atypical_days = points_per_day[
    ~points_per_day.isin(valid_daily_counts)
]


# ============================================================
# FICHE ENTREPRISE
# ============================================================

if company_name or company_siret or selected_location:
    company_parts = []

    if company_name:
        company_parts.append(
            f"<strong>Entreprise :</strong> {company_name}"
        )

    if company_siret:
        company_parts.append(
            f"<strong>SIRET :</strong> {company_siret}"
        )

    if selected_location:
        company_parts.append(
            f"<strong>Adresse :</strong> {selected_location['label']}"
        )

    if advisor_name:
        company_parts.append(
            f"<strong>Conseiller :</strong> {advisor_name}"
        )

    company_parts.append(
        "<strong>Date du diagnostic :</strong> "
        f"{diagnostic_date.strftime('%d/%m/%Y')}"
    )

    st.markdown(
        '<div class="intro-card">'
        '<div class="intro-icon">🏢</div>'
        '<div>'
        + "<br>".join(company_parts)
        + "</div></div>",
        unsafe_allow_html=True,
    )


# ============================================================
# INFORMATIONS DE CONTRÔLE
# ============================================================

info1, info2, info3 = st.columns([2, 1, 1])

with info1:
    st.info(
        f"📅 Données analysées du "
        f"**{filtered_df['Horodate'].min():%d/%m/%Y %H:%M}** "
        f"au **{filtered_df['Horodate'].max():%d/%m/%Y %H:%M}**"
    )

with info2:
    st.info(
        f"⏱ Pas : **{int(time_step.total_seconds() / 60)} min**"
    )

with info3:
    st.info(f"📐 Unité : **{source_unit}**")

st.caption(interpretation)


# ============================================================
# ONGLETS
# ============================================================

(
    tab_dashboard,
    tab_solar,
    tab_tariff,
    tab_financial,
    tab_profiles,
    tab_daily,
    tab_quality,
    tab_export,
) = st.tabs(
    [
        "📊 Tableau de bord",
        "☀️ Analyse solaire",
        "⚡ Analyse tarifaire",
        "💶 Étude financière",
        "🕒 Profils horaires",
        "📅 Consommations journalières",
        "✅ Qualité des données",
        "📥 Export",
    ]
)


# ============================================================
# TABLEAU DE BORD
# ============================================================

with tab_dashboard:
    if pvgis_available:
        render_score_card(cma_score_data)

        with st.expander(
            "Comprendre le calcul de l'indice photovoltaïque CMA"
        ):
            st.markdown(
                f"""
                L'indice est un **repère pédagogique**, et non une note de
                rentabilité ou une validation technique.

                - **Correspondance usages / production PV — 30 %** :
                  {cma_score_data['overlap_score']:.0f}/100
                - **Autoconsommation — 25 %** :
                  {cma_score_data['self_consumption_score']:.0f}/100
                - **Autoproduction — 20 %** :
                  {cma_score_data['self_sufficiency_score']:.0f}/100
                - **Régularité de la consommation — 15 %** :
                  {cma_score_data['regularity_score']:.0f}/100
                - **Potentiel solaire local — 10 %** :
                  {cma_score_data['solar_resource_score']:.0f}/100

                Un score élevé indique qu'il est pertinent d'approfondir le
                projet. Il ne garantit ni sa faisabilité technique, ni sa
                rentabilité économique.
                """
            )

    metric1, metric2, metric3, metric4, metric5 = st.columns(5)

    metric1.metric(
        "Consommation totale",
        f"{format_fr(total_kwh, 0)} kWh",
        help=(
            "Énergie totale consommée sur la période sélectionnée. "
            "Elle permet d'évaluer le volume global des besoins, mais ne "
            "suffit pas à déterminer la puissance photovoltaïque adaptée."
        ),
    )

    metric2.metric(
        "Moyenne journalière",
        f"{format_fr(average_daily_kwh, 1)} kWh",
        help=(
            "Consommation moyenne par journée analysée. Cet indicateur "
            "donne un ordre de grandeur des besoins quotidiens, mais masque "
            "les différences entre jours ouvrés, week-ends et saisons."
        ),
    )

    metric3.metric(
        "Pic de puissance",
        f"{format_fr(maximum_power_kw, 1)} kW",
        help=(
            "Puissance moyenne la plus élevée observée sur un intervalle. "
            "Ce n'est pas la consommation annuelle : il s'agit du niveau "
            "maximal de puissance appelé à un moment donné."
        ),
    )

    metric4.metric(
        "Part pendant le jour",
        (
            f"{format_fr(daylight_share, 1)} %"
            if solar_analysis_available
            else "Adresse requise"
        ),
        help=(
            "Part de la consommation ayant lieu lorsque le soleil est "
            "au-dessus de l'horizon. C'est un premier repère, mais il ne "
            "tient pas compte de l'intensité du soleil. L'indicateur "
            "« consommation pendant la production PV » est plus précis."
        ),
    )

    metric5.metric(
        "Facteur de charge",
        f"{format_fr(load_factor, 1)} %",
        help=(
            "Rapport entre la puissance moyenne et la puissance maximale. "
            "Un facteur élevé traduit généralement une consommation plus "
            "stable ; un facteur faible indique des pointes marquées."
        ),
    )

    chart1, chart2 = st.columns([1.5, 1])

    with chart1:
        fig_monthly = px.bar(
            monthly_df,
            x="Mois_date",
            y="Consommation_kWh",
            title="Consommation mensuelle",
            labels={
                "Mois_date": "Mois",
                "Consommation_kWh": "Consommation (kWh)",
            },
            color_discrete_sequence=[CMA_BLUE],
        )

        fig_monthly.update_layout(
            template="plotly_white",
            showlegend=False,
            hovermode="x unified",
        )

        st.plotly_chart(
            fig_monthly,
            use_container_width=True,
        )

    with chart2:
        fig_solar = go.Figure(
            data=[
                go.Pie(
                    labels=[
                        "Pendant le jour astronomique",
                        "Nuit",
                    ],
                    values=[
                        (
                            daylight_kwh
                            if solar_analysis_available
                            else 0
                        ),
                        (
                            max(total_kwh - daylight_kwh, 0)
                            if solar_analysis_available
                            else total_kwh
                        ),
                    ],
                    hole=0.62,
                    marker=dict(
                        colors=[CMA_RED, "#DDE6EF"]
                    ),
                    textinfo="label+percent",
                )
            ]
        )

        fig_solar.update_layout(
            title="Répartition de la consommation",
            template="plotly_white",
            showlegend=False,
        )

        st.plotly_chart(
            fig_solar,
            use_container_width=True,
        )

    if pvgis_available:
        st.subheader("Bilan énergétique de la simulation photovoltaïque")

        balance_data = pd.DataFrame(
            {
                "Flux": [
                    "Énergie solaire autoconsommée",
                    "Surplus photovoltaïque estimé",
                    "Électricité restant achetée au réseau",
                ],
                "Energie_kWh": [
                    self_consumed_kwh,
                    pv_surplus_kwh,
                    grid_import_kwh,
                ],
            }
        )

        fig_energy_balance = px.bar(
            balance_data,
            x="Flux",
            y="Energie_kWh",
            text_auto=".0f",
            title=(
                "Répartition annuelle des flux d'énergie "
                "pour le scénario étudié"
            ),
            labels={
                "Flux": "",
                "Energie_kWh": "Énergie (kWh)",
            },
            color="Flux",
            color_discrete_map={
                "Énergie solaire autoconsommée": "#E53935",
                "Surplus photovoltaïque estimé": "#F4A261",
                "Électricité restant achetée au réseau": "#17365D",
            },
        )

        fig_energy_balance.update_layout(
            template="plotly_white",
            showlegend=False,
            hovermode="x unified",
        )

        fig_energy_balance.update_traces(
            texttemplate="%{y:,.0f} kWh",
            textposition="outside",
            cliponaxis=False,
        )

        st.plotly_chart(
            fig_energy_balance,
            use_container_width=True,
        )

        st.caption(
            "Ce graphique distingue l'électricité solaire consommée "
            "directement sur place, le surplus potentiel et la part des "
            "besoins qui resterait fournie par le réseau."
        )

    fig_global = px.line(
        hourly_df,
        x="Horodate",
        y="Puissance_kW",
        title="Évolution de la puissance moyenne horaire",
        labels={
            "Horodate": "Date et heure",
            "Puissance_kW": "Puissance moyenne (kW)",
        },
        color_discrete_sequence=[CMA_RED],
    )

    fig_global.update_layout(
        template="plotly_white",
        hovermode="x unified",
    )

    st.plotly_chart(
        fig_global,
        use_container_width=True,
    )



# ============================================================
# ANALYSE SOLAIRE GÉOLOCALISÉE
# ============================================================

with tab_solar:
    st.subheader("Analyse solaire géolocalisée")

    if pvgis_available:
        render_score_card(cma_score_data)

    if not solar_analysis_available:
        st.info(
            "Saisissez puis validez l'adresse précise de l'entreprise "
            "dans le panneau latéral pour calculer les heures de lever "
            "et de coucher du soleil."
        )
    else:
        st.success(
            f"Adresse utilisée : **{selected_location['label']}**  \n"
            f"Coordonnées : **{selected_location['latitude']:.6f}, "
            f"{selected_location['longitude']:.6f}**  \n"
            f"Source : **{selected_location.get('source', 'Géocodage')}**"
        )

        if solar_error:
            st.warning(solar_error)

        if not pd.isna(daylight_share) and (
            daylight_share < 5 or daylight_share > 95
        ):
            st.warning(
                "⚠️ La part de consommation pendant le jour paraît "
                "inhabituelle. Consultez le diagnostic solaire ci-dessous "
                "pour vérifier les horodatages, les coordonnées et les "
                "heures de lever/coucher."
            )

        first_date_row = (
            filtered_df.sort_values("Horodate")
            .dropna(subset=["Lever_soleil", "Coucher_soleil"])
            .iloc[0]
        )
        last_date_row = (
            filtered_df.sort_values("Horodate")
            .dropna(subset=["Lever_soleil", "Coucher_soleil"])
            .iloc[-1]
        )

        s1, s2, s3, s4 = st.columns(4)

        s1.metric(
            "Lever au début de période",
            first_date_row["Lever_soleil"].strftime("%H:%M"),
            help=(
                "Heure astronomique du lever du soleil au premier jour "
                "analysé. La production photovoltaïque débute généralement "
                "progressivement après cette heure."
            ),
        )
        s2.metric(
            "Coucher au début de période",
            first_date_row["Coucher_soleil"].strftime("%H:%M"),
            help=(
                "Heure astronomique du coucher du soleil au premier jour "
                "analysé. La production devient très faible avant d'atteindre "
                "cette heure."
            ),
        )
        s3.metric(
            "Lever en fin de période",
            last_date_row["Lever_soleil"].strftime("%H:%M"),
            help=(
                "Heure du lever du soleil au dernier jour analysé. La "
                "différence avec le début de période illustre la variation "
                "saisonnière de la durée du jour."
            ),
        )
        s4.metric(
            "Coucher en fin de période",
            last_date_row["Coucher_soleil"].strftime("%H:%M"),
            help=(
                "Heure du coucher du soleil au dernier jour analysé. Elle "
                "permet de visualiser l'allongement ou le raccourcissement "
                "des journées sur la période."
            ),
        )

        k1, k2, k3, k4 = st.columns(4)

        k1.metric(
            "Consommation pendant le jour",
            f"{format_fr(daylight_kwh, 0)} kWh",
            f"{format_fr(daylight_share, 1)} %",
            help=(
                "Énergie consommée lorsque le soleil est au-dessus de "
                "l'horizon. Une part élevée est généralement favorable, "
                "mais cet indicateur ne tient pas compte de l'intensité du "
                "rayonnement solaire."
            ),
        )

        k2.metric(
            "Conso. pendant la production PV",
            (
                f"{format_fr(production_period_kwh, 0)} kWh"
                if pvgis_available
                else "PVGIS indisponible"
            ),
            (
                f"{format_fr(production_period_share, 1)} %"
                if pvgis_available
                else None
            ),
            help=(
                "Consommation observée uniquement pendant les intervalles "
                "où l'installation simulée produit effectivement de "
                "l'électricité. Cet indicateur tient compte de la variation "
                "horaire du rayonnement estimé par PVGIS."
            ),
        )

        k3.metric(
            f"Production estimée {pv_peak_kwp:g} kWc",
            (
                f"{format_fr(pvgis_production_kwh, 0)} kWh"
                if pvgis_available
                else "PVGIS indisponible"
            ),
            help=(
                "Énergie photovoltaïque que produirait le scénario étudié "
                "sur la période, selon PVGIS. Le calcul tient compte de la "
                "localisation, de la puissance, de l'orientation, de "
                "l'inclinaison et des pertes renseignées."
            ),
        )

        k4.metric(
            "Taux d'autoproduction estimé",
            (
                f"{format_fr(self_sufficiency_rate, 1)} %"
                if pvgis_available
                else "PVGIS indisponible"
            ),
            help=(
                "Part de la consommation de l'entreprise couverte par "
                "l'électricité solaire autoconsommée. Par exemple, 25 % "
                "signifie qu'environ un quart des besoins serait produit "
                "et consommé sur place."
            ),
        )

        if pvgis_available:
            overlap_label, overlap_color = metric_status(
                production_period_share,
                [
                    (70, "Très bonne correspondance horaire", "#2E8B57"),
                    (50, "Correspondance favorable", "#69A84F"),
                    (30, "Correspondance partielle", "#E0A800"),
                    (0, "Correspondance limitée", "#C0392B"),
                ],
            )
            autocons_label, autocons_color = metric_status(
                self_consumption_rate,
                [
                    (85, "Très forte valorisation sur place", "#2E8B57"),
                    (70, "Bonne valorisation sur place", "#69A84F"),
                    (50, "Valorisation moyenne", "#E0A800"),
                    (0, "Surplus potentiellement important", "#E67E22"),
                ],
            )
            autoprod_label, autoprod_color = metric_status(
                self_sufficiency_rate,
                [
                    (40, "Couverture importante des besoins", "#2E8B57"),
                    (25, "Couverture significative", "#69A84F"),
                    (15, "Couverture modérée", "#E0A800"),
                    (0, "Couverture limitée", "#E67E22"),
                ],
            )

            st.markdown(
                f"""
                <div class="pedagogy-card">
                    <strong>Lecture pédagogique des résultats</strong><br>
                    {render_status_pill(overlap_label, overlap_color)}
                    {render_status_pill(autocons_label, autocons_color)}
                    {render_status_pill(autoprod_label, autoprod_color)}
                    <br><br>
                    Le <strong>taux d'autoconsommation</strong> indique ce que
                    l'entreprise utilise de sa production solaire. Le
                    <strong>taux d'autoproduction</strong> indique la part de
                    ses besoins couverte par cette production. Ces deux taux
                    répondent donc à des questions différentes.
                </div>
                """,
                unsafe_allow_html=True,
            )

            a1, a2 = st.columns(2)

            with a1:
                st.metric(
                    "Énergie PV autoconsommée estimée",
                    f"{format_fr(self_consumed_kwh, 0)} kWh",
                    help=(
                        "Quantité d'électricité solaire utilisée directement "
                        "par le bâtiment. C'est cette énergie qui évite un "
                        "achat équivalent auprès du fournisseur, sous réserve "
                        "des tarifs et conditions du contrat."
                    ),
                )

            with a2:
                st.metric(
                    "Taux d'autoconsommation estimé",
                    f"{format_fr(self_consumption_rate, 1)} %",
                    help=(
                        "Part de la production photovoltaïque consommée "
                        "immédiatement sur place. Un taux élevé limite le "
                        "surplus, mais ne signifie pas forcément que les "
                        "panneaux couvrent une grande part des besoins."
                    ),
                )

            solar_plot = filtered_df[
                [
                    "Horodate",
                    "Puissance_kW",
                    "Production_PV_kW",
                    "Irradiation_Wm2",
                ]
            ].copy()

            fig_power_compare = go.Figure()

            fig_power_compare.add_trace(
                go.Scatter(
                    x=solar_plot["Horodate"],
                    y=solar_plot["Puissance_kW"],
                    name="Consommation",
                    mode="lines",
                    line=dict(color=CMA_BLUE, width=1.5),
                )
            )

            fig_power_compare.add_trace(
                go.Scatter(
                    x=solar_plot["Horodate"],
                    y=solar_plot["Production_PV_kW"],
                    name="Production PV estimée",
                    mode="lines",
                    line=dict(color=CMA_RED, width=1.5),
                )
            )

            fig_power_compare.update_layout(
                title=(
                    "Consommation et production photovoltaïque "
                    "de référence"
                ),
                xaxis_title="Date et heure",
                yaxis_title="Puissance (kW)",
                template="plotly_white",
                hovermode="x unified",
            )

            st.plotly_chart(
                fig_power_compare,
                use_container_width=True,
            )

            fig_irradiation = px.line(
                solar_plot,
                x="Horodate",
                y="Irradiation_Wm2",
                title="Irradiation solaire de référence PVGIS",
                labels={
                    "Horodate": "Date et heure",
                    "Irradiation_Wm2": "Irradiation (W/m²)",
                },
                color_discrete_sequence=[CMA_RED],
            )

            fig_irradiation.update_layout(
                template="plotly_white",
                hovermode="x unified",
            )

            st.plotly_chart(
                fig_irradiation,
                use_container_width=True,
            )

            st.subheader("Bilan des flux d'énergie")

            solar_balance_data = pd.DataFrame(
                {
                    "Flux": [
                        "Solaire autoconsommé",
                        "Surplus photovoltaïque",
                        "Achat au réseau",
                    ],
                    "Energie_kWh": [
                        self_consumed_kwh,
                        pv_surplus_kwh,
                        grid_import_kwh,
                    ],
                }
            )

            fig_solar_balance = go.Figure(
                data=[
                    go.Pie(
                        labels=solar_balance_data["Flux"],
                        values=solar_balance_data["Energie_kWh"],
                        hole=0.55,
                        marker=dict(
                            colors=[
                                CMA_RED,
                                "#F4A261",
                                CMA_BLUE,
                            ]
                        ),
                        textinfo="label+percent",
                        hovertemplate=(
                            "%{label}<br>"
                            "%{value:,.0f} kWh"
                            "<extra></extra>"
                        ),
                    )
                ]
            )

            fig_solar_balance.update_layout(
                title=(
                    "Autoconsommation, surplus et électricité "
                    "achetée au réseau"
                ),
                template="plotly_white",
                showlegend=False,
            )

            st.plotly_chart(
                fig_solar_balance,
                use_container_width=True,
            )

            st.markdown(
                f"""
                **Lecture simple :**

                - **{format_fr(self_consumed_kwh, 0)} kWh** de production
                  solaire pourraient être consommés directement ;
                - **{format_fr(pv_surplus_kwh, 0)} kWh** constitueraient un
                  surplus potentiel ;
                - **{format_fr(grid_import_kwh, 0)} kWh** resteraient à
                  acheter au réseau.
                """
            )

        st.subheader("Diagnostic du calcul solaire")

        d1, d2, d3, d4 = st.columns(4)

        d1.metric(
            "Relevés analysés",
            solar_rows_count,
        )
        d2.metric(
            "Relevés classés en journée",
            solar_day_rows_count,
        )
        d3.metric(
            "Relevés avec lever/coucher",
            solar_event_rows_count,
        )
        d4.metric(
            "Cohérence des 2 méthodes",
            (
                f"{format_fr(solar_coherence_rate, 1)} %"
                if not pd.isna(solar_coherence_rate)
                else "N/D"
            ),
        )

        diagnostic_columns = [
            "Horodate",
            "Horodate_milieu",
            "Hauteur_soleil_deg",
            "Lever_soleil",
            "Coucher_soleil",
            "Soleil_leve",
            "Dans_intervalle_lever_coucher",
            "Controle_solaire_coherent",
            "Energie_kWh",
        ]

        diagnostic_table = filtered_df[
            [
                column
                for column in diagnostic_columns
                if column in filtered_df.columns
            ]
        ].head(48).copy()

        for datetime_column in [
            "Horodate",
            "Horodate_milieu",
            "Lever_soleil",
            "Coucher_soleil",
        ]:
            if datetime_column in diagnostic_table.columns:
                diagnostic_table[datetime_column] = (
                    pd.to_datetime(
                        diagnostic_table[datetime_column],
                        errors="coerce",
                    )
                    .dt.strftime("%d/%m/%Y %H:%M")
                )

        st.dataframe(
            diagnostic_table,
            use_container_width=True,
            height=520,
        )

        st.caption(
            "La colonne « Horodate_milieu » correspond au milieu de "
            "l'intervalle de consommation. La journée est déterminée "
            "principalement à partir de la hauteur apparente du soleil. "
            "La comparaison avec les heures de lever et coucher sert de "
            "contrôle indépendant."
        )

        st.subheader("Lever et coucher du soleil par jour")

        sunrise_table = (
            filtered_df[
                [
                    "Date_solaire",
                    "Lever_soleil",
                    "Midi_solaire",
                    "Coucher_soleil",
                ]
            ]
            .drop_duplicates()
            .sort_values("Date_solaire")
            .copy()
        )

        sunrise_table["Date"] = (
            sunrise_table["Date_solaire"]
            .dt.strftime("%d/%m/%Y")
        )
        sunrise_table["Lever"] = (
            sunrise_table["Lever_soleil"]
            .dt.strftime("%H:%M")
        )
        sunrise_table["Midi solaire"] = (
            sunrise_table["Midi_solaire"]
            .dt.strftime("%H:%M")
        )
        sunrise_table["Coucher"] = (
            sunrise_table["Coucher_soleil"]
            .dt.strftime("%H:%M")
        )
        sunrise_table["Durée du jour"] = (
            sunrise_table["Coucher_soleil"]
            - sunrise_table["Lever_soleil"]
        ).dt.total_seconds() / 3600

        st.dataframe(
            sunrise_table[
                [
                    "Date",
                    "Lever",
                    "Midi solaire",
                    "Coucher",
                    "Durée du jour",
                ]
            ].style.format(
                {"Durée du jour": "{:.2f} h"}
            ),
            use_container_width=True,
            height=440,
        )

        st.caption(
            "Les heures de lever et de coucher sont des calculs "
            "astronomiques précis pour les coordonnées retenues. "
            "Les valeurs PVGIS correspondent à un profil de référence "
            "moyen calculé sur 2020-2023 ; elles ne constituent pas une "
            "mesure météorologique réelle de chaque journée analysée."
        )


# ============================================================
# ANALYSE TARIFAIRE
# ============================================================

with tab_tariff:
    st.subheader("Analyse tarifaire HP / HC et saison haute / saison basse")

    st.markdown(
        f"""
        <div class="score-card" style="--score-color:{tariff_score_data['color']};">
            <div class="score-circle">
                <div class="score-number">{tariff_score_data['score']:.0f}</div>
                <div class="score-total">sur 100</div>
            </div>
            <div>
                <div class="score-title">
                    Indice de potentiel d'optimisation tarifaire CMA
                </div>
                <p class="score-text">
                    {build_tariff_commentary(tariff_score_data)}
                </p>
                {render_status_pill(
                    tariff_score_data['label'],
                    tariff_score_data['color'],
                )}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.warning(
        "Cet indice mesure le potentiel d'analyse et non la qualité du contrat. "
        "Un score élevé signifie qu'il existe davantage de points à approfondir. "
        "Il ne permet pas, à lui seul, de conclure qu'une option tarifaire est "
        "plus économique."
    )

    t1, t2, t3, t4 = st.columns(4)

    t1.metric(
        "HP hiver",
        f"{format_fr(hp_winter_kwh, 0)} kWh",
        help=(
            "Consommation en heures pleines entre le 1er novembre et "
            "le 31 mars, selon les plages d'heures creuses saisies."
        ),
    )
    t2.metric(
        "HC hiver",
        f"{format_fr(hc_winter_kwh, 0)} kWh",
        help=(
            "Consommation en heures creuses entre le 1er novembre et "
            "le 31 mars."
        ),
    )
    t3.metric(
        "HP été",
        f"{format_fr(hp_summer_kwh, 0)} kWh",
        help=(
            "Consommation en heures pleines entre le 1er avril et "
            "le 31 octobre."
        ),
    )
    t4.metric(
        "HC été",
        f"{format_fr(hc_summer_kwh, 0)} kWh",
        help=(
            "Consommation en heures creuses entre le 1er avril et "
            "le 31 octobre."
        ),
    )

    if coverage_ratio < 0.95:
        st.info(
            f"La période analysée couvre environ "
            f"{format_fr(coverage_ratio * 100, 1)} % d'une année. "
            "Les résultats présentés sont ceux de la période disponible "
            "et ne doivent pas être assimilés à une année complète."
        )

    tariff_chart_data = tariff_summary_df.copy()

    tariff_colors = {
        "HP hiver": "#C0392B",
        "HC hiver": "#E67E22",
        "HP été": "#17365D",
        "HC été": "#2E8B57",
    }

    tariff_col1, tariff_col2 = st.columns(2)

    with tariff_col1:
        fig_tariff_pie = px.pie(
            tariff_chart_data,
            names="Categorie_tarifaire",
            values="Consommation_kWh",
            hole=0.52,
            title="Répartition de la consommation par catégorie",
            color="Categorie_tarifaire",
            color_discrete_map=tariff_colors,
        )
        fig_tariff_pie.update_traces(
            textinfo="label+percent",
            hovertemplate=(
                "%{label}<br>"
                "%{value:,.0f} kWh<br>"
                "%{percent}"
                "<extra></extra>"
            ),
        )
        fig_tariff_pie.update_layout(
            template="plotly_white",
            showlegend=False,
        )
        st.plotly_chart(
            fig_tariff_pie,
            use_container_width=True,
        )

    with tariff_col2:
        fig_tariff_bar = px.bar(
            tariff_chart_data,
            x="Categorie_tarifaire",
            y="Consommation_kWh",
            color="Categorie_tarifaire",
            color_discrete_map=tariff_colors,
            text_auto=".0f",
            title="Consommation par catégorie tarifaire",
            labels={
                "Categorie_tarifaire": "",
                "Consommation_kWh": "Consommation (kWh)",
            },
        )
        fig_tariff_bar.update_layout(
            template="plotly_white",
            showlegend=False,
        )
        fig_tariff_bar.update_traces(
            texttemplate="%{y:,.0f} kWh",
            textposition="outside",
            cliponaxis=False,
        )
        st.plotly_chart(
            fig_tariff_bar,
            use_container_width=True,
        )

    hp_total = hp_winter_kwh + hp_summer_kwh
    hc_total = hc_winter_kwh + hc_summer_kwh
    winter_total = hp_winter_kwh + hc_winter_kwh
    summer_total = hp_summer_kwh + hc_summer_kwh

    s1, s2, s3, s4 = st.columns(4)
    s1.metric(
        "Part totale en HP",
        f"{format_fr(tariff_score_data['hp_share'], 1)} %",
        help=(
            "Part de la consommation totale située en heures pleines. "
            "Une part élevée peut inviter à rechercher les usages décalables, "
            "mais le déplacement doit rester compatible avec l'activité."
        ),
    )
    s2.metric(
        "Part totale en HC",
        f"{format_fr(tariff_score_data['hc_share'], 1)} %",
        help=(
            "Part de la consommation totale située dans les plages d'heures "
            "creuses renseignées."
        ),
    )
    s3.metric(
        "Consommation hiver",
        f"{format_fr(winter_total, 0)} kWh",
        help="Total saison haute, du 1er novembre au 31 mars.",
    )
    s4.metric(
        "Consommation été",
        f"{format_fr(summer_total, 0)} kWh",
        help="Total saison basse, du 1er avril au 31 octobre.",
    )

    with st.expander(
        "Comprendre l'indice d'optimisation tarifaire CMA"
    ):
        st.markdown(
            f"""
            Cet indice est un **repère pédagogique de potentiel d'analyse**.

            - **Part consommée en heures pleines — 50 %** :
              {tariff_score_data['hp_opportunity_score']:.0f}/100
            - **Écart entre saison haute et saison basse — 25 %** :
              {tariff_score_data['seasonality_score']:.0f}/100
            - **Variabilité des consommations — 15 %** :
              {tariff_score_data['variability_score']:.0f}/100
            - **Complétude de la période — 10 %** :
              {tariff_score_data['coverage_score']:.0f}/100

            Un score élevé signifie qu'une analyse complémentaire du contrat,
            des prix et des usages décalables peut être pertinente. Le score
            ne compare pas les offres commerciales des fournisseurs.
            """
        )

    st.subheader("Tableau récapitulatif")

    tariff_display = tariff_summary_df.copy()
    tariff_display.columns = [
        "Catégorie",
        "Consommation (kWh)",
        "Nombre d'intervalles",
        "Part (%)",
    ]

    st.dataframe(
        tariff_display.style.format(
            {
                "Consommation (kWh)": "{:.2f}",
                "Part (%)": "{:.1f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Lecture pédagogique")

    st.markdown(
        f"""
        <div class="pedagogy-card">
            <strong>Ce que montre l'analyse</strong><br><br>
            {build_tariff_commentary(tariff_score_data)}<br><br>
            Les plages d'heures creuses utilisées sont celles saisies dans
            le panneau latéral. Toutes les autres heures sont considérées
            comme des heures pleines. Avant toute recommandation, il faut
            comparer les résultats avec les prix réels du contrat et vérifier
            quels usages peuvent réellement être déplacés.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# ÉTUDE FINANCIÈRE
# ============================================================

with tab_financial:
    st.subheader("Étude financière indicative du projet")

    st.warning(
        "Cette simulation fournit des ordres de grandeur HT. "
        "Elle ne remplace pas les devis, l'étude structure, la proposition "
        "de raccordement Enedis, l'analyse fiscale ou le plan de financement."
    )

    if not pvgis_available:
        st.info(
            "L'adresse et les données PVGIS sont nécessaires pour valoriser "
            "l'autoconsommation et le surplus. Les coûts d'investissement "
            "restent néanmoins consultables."
        )

    st.markdown(
        f"""
        <div class="score-card" style="--score-color:{business_assistant['color']};">
            <div class="score-circle">
                <div class="score-number">{cma_score_data['score']:.0f}</div>
                <div class="score-total">score PV / 100</div>
            </div>
            <div>
                <div class="score-title">
                    Assistant CMA — {business_assistant['headline']}
                </div>
                <p class="score-text">{business_assistant['conclusion']}</p>
                {render_status_pill(
                    business_assistant['status'],
                    business_assistant['color'],
                )}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    assistant_col1, assistant_col2 = st.columns(2)

    with assistant_col1:
        st.markdown(
            f"""
            <div class="pedagogy-card">
                <strong>✅ Points favorables</strong><br><br>
                {assistant_html_list(
                    business_assistant['strengths'],
                    '•',
                )}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with assistant_col2:
        st.markdown(
            f"""
            <div class="pedagogy-card">
                <strong>⚠️ Points de vigilance</strong><br><br>
                {assistant_html_list(
                    business_assistant['vigilance'],
                    '•',
                )}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with st.expander("🧭 Prochaines étapes recommandées"):
        for step_number, step_text in enumerate(
            business_assistant["next_steps"],
            start=1,
        ):
            st.markdown(f"**{step_number}.** {step_text}")

    f1, f2, f3, f4 = st.columns(4)

    f1.metric(
        "Investissement brut",
        f"{format_fr(investment_data['gross_total'], 0)} € HT",
        help="Somme des équipements, fixation, études, toiture, raccordement et autres coûts.",
    )
    f2.metric(
        "Investissement net",
        f"{format_fr(investment_data['net_total'], 0)} € HT",
        help="Investissement brut diminué des aides ou subventions saisies.",
    )
    f3.metric(
        "Gain net estimé en année 1",
        f"{format_fr(financial_projection['annual_net_gain_year_1'], 0)} €",
        help="Économies d'autoconsommation + revenus du surplus - charges annuelles.",
    )
    f4.metric(
        "Temps de retour simple",
        (
            f"{format_fr(financial_projection['payback_year'], 1)} ans"
            if not pd.isna(financial_projection["payback_year"])
            else "Au-delà de l'horizon"
        ),
    )

    f5, f6, f7, f8 = st.columns(4)

    f5.metric(
        f"VAN à {financial_horizon_years} ans",
        f"{format_fr(financial_projection['npv'], 0)} €",
        help="Valeur actuelle nette calculée avec le taux d'actualisation renseigné.",
    )
    f6.metric(
        "TRI estimé",
        (
            f"{format_fr(financial_projection['irr'] * 100, 1)} %"
            if not pd.isna(financial_projection["irr"])
            else "Non calculable"
        ),
    )
    f7.metric(
        f"Gain net cumulé à {financial_horizon_years} ans",
        f"{format_fr(financial_projection['total_net_gain'], 0)} €",
    )
    f8.metric(
        "Charges annuelles initiales",
        f"{format_fr(operating_cost_data['total'], 0)} € / an",
    )

    st.subheader("Décomposition de l'investissement")

    investment_breakdown = pd.DataFrame(
        {
            "Poste": [
                "Modules, onduleur, câblage et pose",
                "Système de fixation",
                "Surcoût ERP / ICPE",
                "Étude structure",
                "Rénovation de couverture",
                "Désamiantage",
                "Raccordement",
                "Autres coûts",
            ],
            "Montant_EUR": [
                investment_data["equipment_cost"],
                investment_data["fixing_cost"],
                investment_data["erp_surcharge_cost"],
                investment_data["structural_study_cost"],
                investment_data["roof_cost"],
                investment_data["asbestos_cost"],
                investment_data["connection_cost"],
                investment_data["other_investment_costs"],
            ],
        }
    )
    investment_breakdown = investment_breakdown[
        investment_breakdown["Montant_EUR"] > 0
    ]

    # Affichage en pleine largeur : plus robuste que deux colonnes lorsque
    # le navigateur applique un zoom ou que la fenêtre est étroite.
    chart_height = max(
        380,
        105 + 58 * len(investment_breakdown),
    )

    fig_investment = px.bar(
        investment_breakdown,
        x="Montant_EUR",
        y="Poste",
        orientation="h",
        text_auto=".0f",
        title="Répartition du coût d'investissement",
        labels={
            "Montant_EUR": "Montant (€ HT)",
            "Poste": "",
        },
    )
    fig_investment.update_traces(
        texttemplate="%{x:,.0f} €",
        textposition="outside",
        cliponaxis=False,
    )
    fig_investment.update_layout(
        height=chart_height,
        margin=dict(
            l=210,
            r=110,
            t=75,
            b=65,
        ),
        yaxis=dict(
            automargin=True,
            categoryorder="total ascending",
        ),
        xaxis=dict(
            automargin=True,
            rangemode="tozero",
        ),
    )

    st.plotly_chart(
        fig_investment,
        use_container_width=True,
    )

    investment_table = investment_breakdown.copy()
    investment_table.columns = ["Poste", "Montant (€ HT)"]
    total_row = pd.DataFrame(
        {
            "Poste": [
                "TOTAL BRUT",
                "Aides déduites",
                "TOTAL NET",
            ],
            "Montant (€ HT)": [
                investment_data["gross_total"],
                -investment_data["grant_amount"],
                investment_data["net_total"],
            ],
        }
    )
    investment_table = pd.concat(
        [investment_table, total_row],
        ignore_index=True,
    )

    st.dataframe(
        investment_table.style.format(
            {"Montant (€ HT)": "{:,.0f} €"}
        ),
        use_container_width=True,
        hide_index=True,
        height=min(
            520,
            42 + 35 * len(investment_table),
        ),
    )

    # Séparation explicite pour empêcher tout chevauchement avec le bloc suivant.
    st.markdown(
        '<div style="height:22px"></div>',
        unsafe_allow_html=True,
    )

    st.subheader("Raccordement indicatif")

    connection_table = pd.DataFrame(
        {
            "Poste": [
                "Ouvrages publics avant réfaction",
                "Réfaction Enedis estimative",
                "Ouvrages publics après réfaction",
                "Tranchée privée",
                "Poste privé HTA/BT",
                "Cellule de découplage",
                "Total raccordement",
            ],
            "Montant (€ HT)": [
                connection_data["public_gross"],
                -connection_data["enedis_reduction"],
                connection_data["public_net"],
                connection_data["private_trench"],
                connection_data["private_post"],
                connection_data["decoupling_cell"],
                connection_data["total"],
            ],
        }
    )

    st.dataframe(
        connection_table.style.format(
            {"Montant (€ HT)": "{:,.0f} €"}
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Valorisation annuelle de l'énergie")

    e1, e2, e3, e4 = st.columns(4)

    e1.metric(
        "Facture annuelle de référence",
        f"{format_fr(energy_value_data['annual_energy_bill'], 0)} € HT",
        help="Estimation annualisée à partir de la courbe de charge et des prix saisis, abonnement inclus.",
    )
    e2.metric(
        "Économie d'autoconsommation",
        f"{format_fr(energy_value_data['annual_self_consumption_saving'], 0)} € / an",
    )
    e3.metric(
        "Revenu du surplus",
        f"{format_fr(energy_value_data['annual_surplus_revenue'], 0)} € / an",
    )
    e4.metric(
        "Charges annuelles",
        f"{format_fr(operating_cost_data['total'], 0)} € / an",
    )

    operating_table = pd.DataFrame(
        {
            "Charge annuelle": [
                "Assurance",
                "Suivi et maintenance",
                "Provision onduleurs",
                "TURPE",
                "IFER",
                "Autres charges",
                "TOTAL",
            ],
            "Montant (€ HT/an)": [
                operating_cost_data["insurance"],
                operating_cost_data["maintenance"],
                operating_cost_data["inverter_provision"],
                operating_cost_data["turpe"],
                operating_cost_data["ifer"],
                operating_cost_data["other"],
                operating_cost_data["total"],
            ],
        }
    )

    st.dataframe(
        operating_table.style.format(
            {"Montant (€ HT/an)": "{:,.0f} €"}
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Projection sur la durée du projet")

    projection_df = financial_projection["table"].copy()

    fig_cashflow = go.Figure()
    fig_cashflow.add_trace(
        go.Bar(
            x=projection_df["Année"],
            y=projection_df["Flux net (€)"],
            name="Flux net annuel",
        )
    )
    fig_cashflow.add_trace(
        go.Scatter(
            x=projection_df["Année"],
            y=projection_df["Cumul net (€)"],
            name="Cumul net",
            mode="lines+markers",
            yaxis="y2",
        )
    )
    fig_cashflow.add_hline(
        y=0,
        line_dash="dash",
    )
    fig_cashflow.update_layout(
        title="Flux financiers et cumul du projet",
        xaxis_title="Année",
        yaxis_title="Flux annuel (€)",
        yaxis2=dict(
            title="Cumul (€)",
            overlaying="y",
            side="right",
            showgrid=False,
        ),
        legend=dict(orientation="h"),
    )
    st.plotly_chart(fig_cashflow, use_container_width=True)

    with st.expander("Afficher le détail annuel"):
        st.dataframe(
            projection_df.style.format(
                {
                    "Économie autoconsommation (€)": "{:,.0f}",
                    "Revenu surplus (€)": "{:,.0f}",
                    "Charges annuelles (€)": "{:,.0f}",
                    "Flux net (€)": "{:,.0f}",
                    "Flux actualisé (€)": "{:,.0f}",
                    "Cumul net (€)": "{:,.0f}",
                    "Cumul actualisé (€)": "{:,.0f}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("Hypothèses utilisées")

    hypothesis_table = pd.DataFrame(
        {
            "Hypothèse": [
                "Type de tarif électrique",
                "Prix de vente du surplus",
                "Puissance étudiée",
                "Fixation",
                "Coût équipements et pose",
                "Coût fixation",
                "Durée de projection",
                "Hausse prix électricité",
                "Dégradation production",
                "Taux d'actualisation",
            ],
            "Valeur": [
                electricity_tariff_type,
                f"{surplus_sale_price_eur_kwh:.4f} €/kWh HT",
                f"{pv_peak_kwp:g} kWc",
                fixing_type,
                f"{investment_data['equipment_rate']:.2f} €/Wc",
                f"{investment_data['fixing_rate']:.2f} €/Wc",
                f"{financial_horizon_years} ans",
                f"{electricity_price_increase_percent:.1f} %/an",
                f"{production_degradation_percent:.1f} %/an",
                f"{discount_rate_percent:.1f} %",
            ],
        }
    )
    st.dataframe(
        hypothesis_table,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# PROFILS HORAIRES
# ============================================================

with tab_profiles:
    st.subheader(
        "Consommation horaire moyenne selon le jour de la semaine"
    )

    col_table, col_chart = st.columns([1, 1.35])

    with col_table:
        display_matrix = weekday_hour_matrix.copy()
        display_matrix.index = [
            f"De {hour:02d}:00 à {(hour + 1) % 24:02d}:00"
            for hour in display_matrix.index
        ]
        display_matrix.index.name = "Plage horaire"

        st.dataframe(
            make_colored_style(display_matrix),
            use_container_width=True,
            height=810,
        )

    with col_chart:
        profile_long = (
            weekday_hour_matrix
            .reset_index()
            .melt(
                id_vars="Heure",
                var_name="Jour",
                value_name="Puissance_kW",
            )
            .dropna()
        )

        profile_long["Jour"] = pd.Categorical(
            profile_long["Jour"],
            categories=WEEKDAY_ORDER,
            ordered=True,
        )

        fig_profiles = px.line(
            profile_long,
            x="Heure",
            y="Puissance_kW",
            color="Jour",
            markers=True,
            title=(
                "Profil horaire moyen selon le jour "
                "de la semaine"
            ),
            labels={
                "Heure": "Heure",
                "Puissance_kW": "Puissance moyenne (kW)",
            },
        )

        fig_profiles.update_layout(
            template="plotly_white",
            hovermode="x unified",
            legend_title_text="Jour",
            height=810,
        )

        fig_profiles.update_xaxes(dtick=1)

        st.plotly_chart(
            fig_profiles,
            use_container_width=True,
        )

    st.subheader("Carte thermique interactive")

    heatmap = go.Figure(
        data=go.Heatmap(
            z=weekday_hour_matrix.values,
            x=weekday_hour_matrix.columns,
            y=[
                f"{hour:02d}:00"
                for hour in weekday_hour_matrix.index
            ],
            colorscale=[
                [0.00, "#63BE7B"],
                [0.30, "#A9D26D"],
                [0.50, "#FFEB84"],
                [0.72, "#F6B26B"],
                [1.00, "#F8696B"],
            ],
            colorbar=dict(title="kW"),
            hovertemplate=(
                "Jour : %{x}<br>"
                "Heure : %{y}<br>"
                "Puissance moyenne : %{z:.2f} kW"
                "<extra></extra>"
            ),
        )
    )

    heatmap.update_layout(
        template="plotly_white",
        xaxis_title="Jour de la semaine",
        yaxis_title="Heure",
        height=720,
    )

    heatmap.update_yaxes(autorange="reversed")

    st.plotly_chart(
        heatmap,
        use_container_width=True,
    )


# ============================================================
# CONSOMMATIONS JOURNALIÈRES
# ============================================================

with tab_daily:
    daily1, daily2, daily3 = st.columns(3)

    daily1.metric(
        "Moyenne journalière",
        f"{format_fr(average_daily_kwh, 1)} kWh",
    )

    daily2.metric(
        "Médiane journalière",
        f"{format_fr(median_daily_kwh, 1)} kWh",
    )

    daily3.metric(
        "Maximum journalier",
        f"{format_fr(daily_df['Consommation_kWh'].max(), 1)} kWh",
    )

    fig_daily = px.bar(
        daily_df,
        x="Date",
        y="Consommation_kWh",
        color="Jour",
        category_orders={
            "Jour": WEEKDAY_ORDER,
        },
        title="Consommation quotidienne",
        labels={
            "Date": "Date",
            "Consommation_kWh": "Consommation (kWh)",
        },
    )

    fig_daily.update_layout(
        template="plotly_white",
        hovermode="x unified",
        legend_title_text="Jour",
    )

    st.plotly_chart(
        fig_daily,
        use_container_width=True,
    )

    calendar_years = sorted(
        daily_df["Année"].unique().tolist()
    )

    calendar_year = st.selectbox(
        "Année du calendrier thermique",
        calendar_years,
        index=len(calendar_years) - 1,
    )

    calendar_matrix = build_daily_calendar(
        daily_df,
        calendar_year,
    )

    if not calendar_matrix.empty:
        calendar_heatmap = go.Figure(
            data=go.Heatmap(
                z=calendar_matrix.values,
                x=WEEKDAY_ORDER,
                y=[
                    f"Semaine {int(week)}"
                    for week in calendar_matrix.index
                ],
                colorscale=[
                    [0.00, "#63BE7B"],
                    [0.30, "#A9D26D"],
                    [0.50, "#FFEB84"],
                    [0.72, "#F6B26B"],
                    [1.00, "#F8696B"],
                ],
                colorbar=dict(title="kWh"),
                hovertemplate=(
                    "Jour : %{x}<br>"
                    "%{y}<br>"
                    "Consommation : %{z:.1f} kWh"
                    "<extra></extra>"
                ),
            )
        )

        calendar_heatmap.update_layout(
            title=(
                "Calendrier des consommations journalières "
                f"– {calendar_year}"
            ),
            template="plotly_white",
            height=920,
            xaxis_title="Jour de la semaine",
            yaxis_title="Semaine",
        )

        calendar_heatmap.update_yaxes(
            autorange="reversed"
        )

        st.plotly_chart(
            calendar_heatmap,
            use_container_width=True,
        )

    st.subheader("Tableau détaillé")

    daily_display = daily_df[
        [
            "Date",
            "Jour",
            "Consommation_kWh",
            "Puissance_moyenne_kW",
            "Puissance_max_kW",
        ]
    ].copy()

    daily_display["Date"] = (
        daily_display["Date"]
        .dt.strftime("%d/%m/%Y")
    )

    daily_style = (
        daily_display.style
        .format(
            {
                "Consommation_kWh": "{:.2f}",
                "Puissance_moyenne_kW": "{:.2f}",
                "Puissance_max_kW": "{:.2f}",
            }
        )
        .background_gradient(
            subset=["Consommation_kWh"],
            cmap="RdYlGn_r",
        )
    )

    st.dataframe(
        daily_style,
        use_container_width=True,
        height=520,
    )


# ============================================================
# QUALITÉ DES DONNÉES
# ============================================================

with tab_quality:
    quality1, quality2, quality3, quality4 = st.columns(4)

    quality1.metric(
        "Lignes exploitables",
        f"{len(filtered_df):,}".replace(",", " "),
    )

    quality2.metric(
        "Horodatages en doublon",
        duplicate_count,
    )

    quality3.metric(
        "Jours atypiques",
        len(atypical_days),
    )

    quality4.metric(
        "Points théoriques/jour",
        expected_points_per_day,
    )

    quality_df = (
        points_per_day
        .rename("Nombre_points")
        .reset_index()
    )

    quality_df.columns = [
        "Date",
        "Nombre_points",
    ]

    quality_df["Statut"] = np.where(
        quality_df["Nombre_points"].isin(
            valid_daily_counts
        ),
        "Cohérent",
        "À contrôler",
    )

    st.dataframe(
        quality_df,
        use_container_width=True,
        height=480,
    )

    if atypical_days.empty:
        st.success(
            "Aucune journée anormale détectée, "
            "hors changements d'heure possibles."
        )
    else:
        st.warning(
            "Certaines journées comportent un nombre "
            "inhabituel de relevés."
        )

        st.dataframe(
            atypical_days
            .rename("Nombre de relevés")
            .to_frame(),
            use_container_width=True,
        )


# ============================================================
# EXPORT
# ============================================================

with tab_export:
    st.subheader("Exporter l'analyse")

    summary_df = pd.DataFrame(
        {
            "Indicateur": [
                "Nom de l'entreprise",
                "SIRET",
                "Conseiller CMA",
                "Date du diagnostic",
                "Fichier source",
                "Début de période",
                "Fin de période",
                "Pas de temps source",
                "Pas de temps après traitement",
                "Unité source",
                "Consommation totale (kWh)",
                "Moyenne journalière (kWh)",
                "Médiane journalière (kWh)",
                "Pic de puissance (kW)",
                "Adresse de l'entreprise",
                "Latitude",
                "Longitude",
                "Part de consommation pendant le jour (%)",
                "Part de consommation pendant la production PV (%)",
                "Puissance photovoltaïque étudiée (kWc)",
                "Production PVGIS estimée (kWh)",
                "Énergie solaire autoconsommée estimée (kWh)",
                "Surplus photovoltaïque estimé (kWh)",
                "Électricité restant achetée au réseau (kWh)",
                "Taux d'autoconsommation estimé (%)",
                "Taux d'autoproduction estimé (%)",
                "Indice photovoltaïque CMA (/100)",
                "Appréciation de l'indice CMA",
                "Score correspondance usages / production (/100)",
                "Score autoconsommation (/100)",
                "Score autoproduction (/100)",
                "Score régularité des consommations (/100)",
                "Score potentiel solaire local (/100)",
                "Productible estimé (kWh/kWc)",
                "HP hiver (kWh)",
                "HC hiver (kWh)",
                "HP été (kWh)",
                "HC été (kWh)",
                "Part totale HP (%)",
                "Part totale HC (%)",
                "Indice optimisation tarifaire CMA (/100)",
                "Appréciation indice tarifaire",
                "Investissement brut estimé (€ HT)",
                "Investissement net estimé (€ HT)",
                "Charges annuelles estimées (€ HT/an)",
                "Économie autoconsommation annuelle (€ HT)",
                "Revenu surplus annuel (€ HT)",
                "Gain net année 1 (€ HT)",
                "Temps de retour simple (années)",
                "VAN du projet (€)",
                "TRI estimé (%)",
                "Gain net cumulé sur l'horizon (€)",
                "Synthèse assistant CMA",
                "Statut assistant CMA",
                "Facteur de charge (%)",
                "Horodatages en doublon",
                "Jours atypiques",
            ],
            "Valeur": [
                company_name,
                company_siret,
                advisor_name,
                diagnostic_date.strftime("%d/%m/%Y"),
                uploaded_file.name,
                filtered_df["Horodate"].min(),
                filtered_df["Horodate"].max(),
                str(time_step),
                "01:00:00",
                source_unit,
                total_kwh,
                average_daily_kwh,
                median_daily_kwh,
                maximum_power_kw,
                (
                    selected_location["label"]
                    if selected_location
                    else ""
                ),
                (
                    selected_location["latitude"]
                    if selected_location
                    else ""
                ),
                (
                    selected_location["longitude"]
                    if selected_location
                    else ""
                ),
                (
                    daylight_share
                    if solar_analysis_available
                    else ""
                ),
                (
                    production_period_share
                    if pvgis_available
                    else ""
                ),
                pv_peak_kwp,
                (
                    pvgis_production_kwh
                    if pvgis_available
                    else ""
                ),
                (
                    self_consumed_kwh
                    if pvgis_available
                    else ""
                ),
                (
                    pv_surplus_kwh
                    if pvgis_available
                    else ""
                ),
                (
                    grid_import_kwh
                    if pvgis_available
                    else ""
                ),
                (
                    self_consumption_rate
                    if pvgis_available
                    else ""
                ),
                (
                    self_sufficiency_rate
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["score"]
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["label"]
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["overlap_score"]
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["self_consumption_score"]
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["self_sufficiency_score"]
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["regularity_score"]
                    if pvgis_available
                    else ""
                ),
                (
                    cma_score_data["solar_resource_score"]
                    if pvgis_available
                    else ""
                ),
                (
                    annual_yield_kwh_per_kwp
                    if pvgis_available
                    else ""
                ),
                hp_winter_kwh,
                hc_winter_kwh,
                hp_summer_kwh,
                hc_summer_kwh,
                tariff_score_data["hp_share"],
                tariff_score_data["hc_share"],
                tariff_score_data["score"],
                tariff_score_data["label"],
                investment_data["gross_total"],
                investment_data["net_total"],
                operating_cost_data["total"],
                energy_value_data["annual_self_consumption_saving"],
                energy_value_data["annual_surplus_revenue"],
                financial_projection["annual_net_gain_year_1"],
                (
                    financial_projection["payback_year"]
                    if not pd.isna(financial_projection["payback_year"])
                    else ""
                ),
                financial_projection["npv"],
                (
                    financial_projection["irr"] * 100
                    if not pd.isna(financial_projection["irr"])
                    else ""
                ),
                financial_projection["total_net_gain"],
                business_assistant["conclusion"],
                business_assistant["status"],
                load_factor,
                duplicate_count,
                len(atypical_days),
            ],
        }
    )

    # Export principal obligatoirement normalisé à un pas horaire.
    # Les relevés de 30 minutes sont agrégés dans hourly_df :
    # - moyenne pour une unité de puissance (W / kW) ;
    # - somme pour une unité d'énergie (Wh / kWh).
    hourly_standardized_export = hourly_df.copy()

    if solar_analysis_available:
        solar_hourly_columns = [
            "Horodate",
            "Lever_soleil",
            "Coucher_soleil",
            "Hauteur_soleil_deg",
        ]

        if pvgis_available:
            solar_hourly_columns += [
                "Irradiation_Wm2",
                "Production_PV_kW",
                "Production_PV_kWh",
                "Autoconsommation_estimee_kWh",
            ]

        solar_hourly_export = filtered_df.copy()
        solar_hourly_export["Horodate_heure"] = (
            solar_hourly_export["Horodate"].dt.ceil("h")
        )

        aggregations = {
            "Lever_soleil": "first",
            "Coucher_soleil": "first",
            "Hauteur_soleil_deg": "mean",
        }

        if pvgis_available:
            aggregations.update(
                {
                    "Irradiation_Wm2": "mean",
                    "Production_PV_kW": "mean",
                    "Production_PV_kWh": "sum",
                    "Autoconsommation_estimee_kWh": "sum",
                }
            )

        solar_hourly_export = (
            solar_hourly_export.groupby(
                "Horodate_heure",
                as_index=False,
            )
            .agg(aggregations)
            .rename(
                columns={"Horodate_heure": "Horodate"}
            )
        )

        hourly_standardized_export = (
            hourly_standardized_export.merge(
                solar_hourly_export,
                on="Horodate",
                how="left",
            )
        )

    normalized_unit = str(source_unit).lower().replace(" ", "")

    if normalized_unit in {"w", "watt", "watts"}:
        hourly_standardized_export["Valeur"] = (
            hourly_standardized_export["Puissance_kW"] * 1000
        )
        export_unit = "W"
    elif normalized_unit in {"kw", "kilowatt", "kilowatts"}:
        hourly_standardized_export["Valeur"] = (
            hourly_standardized_export["Puissance_kW"]
        )
        export_unit = "kW"
    elif normalized_unit in {"kwh", "kw.h"}:
        hourly_standardized_export["Valeur"] = (
            hourly_standardized_export["Energie_kWh"]
        )
        export_unit = "kWh"
    else:
        hourly_standardized_export["Valeur"] = (
            hourly_standardized_export["Energie_kWh"] * 1000
        )
        export_unit = "Wh"

    hourly_standardized_export.insert(0, "Unité", export_unit)
    hourly_standardized_export["Pas"] = "PT60M"
    hourly_standardized_export["Horodate"] = (
        hourly_standardized_export["Horodate"]
        .dt.strftime("%d/%m/%Y %H:%M:%S")
    )
    export_columns = [
        "Unité",
        "Horodate",
        "Valeur",
        "Pas",
    ]

    optional_solar_columns = [
        "Lever_soleil",
        "Coucher_soleil",
        "Hauteur_soleil_deg",
        "Irradiation_Wm2",
        "Production_PV_kW",
        "Production_PV_kWh",
        "Autoconsommation_estimee_kWh",
    ]

    export_columns += [
        column
        for column in optional_solar_columns
        if column in hourly_standardized_export.columns
    ]

    hourly_standardized_export = hourly_standardized_export[
        export_columns
    ]

    hourly_export = hourly_df.copy()
    hourly_export["Horodate"] = (
        hourly_export["Horodate"]
        .dt.strftime("%d/%m/%Y %H:%M:%S")
    )

    daily_export = daily_df.copy()
    daily_export["Date"] = (
        daily_export["Date"]
        .dt.strftime("%d/%m/%Y")
    )

    monthly_export = monthly_df.copy()
    monthly_export["Mois_date"] = (
        monthly_export["Mois_date"]
        .dt.strftime("%m/%Y")
    )

    tariff_summary_export = tariff_summary_df.copy()

    tariff_detail_export = filtered_df[
        [
            "Horodate",
            "Horodate_tarif",
            "Energie_kWh",
            "Saison_tarifaire",
            "Plage_tarifaire",
            "Categorie_tarifaire",
        ]
    ].copy()

    tariff_detail_export["Horodate"] = (
        tariff_detail_export["Horodate"]
        .dt.strftime("%d/%m/%Y %H:%M:%S")
    )
    tariff_detail_export["Horodate_tarif"] = (
        tariff_detail_export["Horodate_tarif"]
        .dt.strftime("%d/%m/%Y %H:%M:%S")
    )

    financial_summary_export = pd.DataFrame(
        {
            "Indicateur": [
                "Puissance étudiée (kWc)",
                "Type de fixation",
                "Coût équipements + pose (€/Wc)",
                "Coût fixation (€/Wc)",
                "Investissement brut (€ HT)",
                "Aides déduites (€)",
                "Investissement net (€ HT)",
                "Raccordement (€ HT)",
                "Charges annuelles (€ HT/an)",
                "Facture annuelle de référence (€ HT)",
                "Économie autoconsommation (€ HT/an)",
                "Revenu surplus (€ HT/an)",
                "Gain net année 1 (€ HT)",
                "Temps de retour simple (années)",
                "VAN (€)",
                "TRI (%)",
                "Gain net cumulé (€)",
            ],
            "Valeur": [
                pv_peak_kwp,
                fixing_type,
                investment_data["equipment_rate"],
                investment_data["fixing_rate"],
                investment_data["gross_total"],
                investment_data["grant_amount"],
                investment_data["net_total"],
                connection_data["total"],
                operating_cost_data["total"],
                energy_value_data["annual_energy_bill"],
                energy_value_data["annual_self_consumption_saving"],
                energy_value_data["annual_surplus_revenue"],
                financial_projection["annual_net_gain_year_1"],
                (
                    financial_projection["payback_year"]
                    if not pd.isna(financial_projection["payback_year"])
                    else ""
                ),
                financial_projection["npv"],
                (
                    financial_projection["irr"] * 100
                    if not pd.isna(financial_projection["irr"])
                    else ""
                ),
                financial_projection["total_net_gain"],
            ],
        }
    )

    excel_bytes = make_excel_export(
        hourly_standardized_export,
        hourly_export,
        daily_export,
        monthly_export,
        weekday_hour_matrix.round(3),
        tariff_summary_export,
        tariff_detail_export,
        financial_summary_export,
        financial_projection["table"],
        summary_df,
    )

    export1, export2 = st.columns(2)

    logo_candidates = [
        Path("logo_cma.png"),
        Path("logo_cma.jpg"),
        Path("assets/logo_cma.png"),
        Path("assets/logo_cma.jpg"),
    ]
    report_logo_path = next(
        (path for path in logo_candidates if path.exists()),
        None,
    )

    pdf_ready = bool(
        solar_analysis_available
        and pvgis_available
        and selected_location is not None
    )

    if pdf_ready:
        try:
            pdf_report_bytes = create_cma_pdf_report(
                company_name=company_name,
                company_siret=company_siret,
                advisor_name=advisor_name,
                diagnostic_date=diagnostic_date,
                address_label=selected_location["label"],
                latitude=selected_location["latitude"],
                longitude=selected_location["longitude"],
                source_filename=uploaded_file.name,
                period_start=filtered_df["Horodate"].min(),
                period_end=filtered_df["Horodate"].max(),
                source_unit=source_unit,
                time_step=time_step,
                total_kwh=total_kwh,
                average_daily_kwh=average_daily_kwh,
                maximum_power_kw=maximum_power_kw,
                daylight_share=daylight_share,
                production_period_share=production_period_share,
                pv_peak_kwp=pv_peak_kwp,
                pv_tilt=pv_tilt,
                orientation_label=orientation_label,
                pv_losses=pv_losses,
                pvgis_production_kwh=pvgis_production_kwh,
                self_consumed_kwh=self_consumed_kwh,
                self_consumption_rate=self_consumption_rate,
                self_sufficiency_rate=self_sufficiency_rate,
                cma_score_data=cma_score_data,
                annual_yield_kwh_per_kwp=annual_yield_kwh_per_kwp,
                tariff_summary_df=tariff_summary_df,
                tariff_score_data=tariff_score_data,
                hc_ranges=hc_ranges,
                investment_data=investment_data,
                connection_data=connection_data,
                operating_cost_data=operating_cost_data,
                energy_value_data=energy_value_data,
                financial_projection=financial_projection,
                business_assistant=business_assistant,
                financial_horizon_years=financial_horizon_years,
                electricity_tariff_type=electricity_tariff_type,
                surplus_sale_price_eur_kwh=surplus_sale_price_eur_kwh,
                electricity_price_increase_percent=(
                    electricity_price_increase_percent
                ),
                production_degradation_percent=(
                    production_degradation_percent
                ),
                discount_rate_percent=discount_rate_percent,
                monthly_df=monthly_df,
                weekday_hour_matrix=weekday_hour_matrix,
                hourly_df=hourly_df,
                filtered_df=filtered_df,
                logo_path=report_logo_path,
            )
        except Exception as exc:
            pdf_report_bytes = None
            st.error(
                "Le rapport PDF n'a pas pu être généré. "
                f"Détail : {exc}"
            )
    else:
        pdf_report_bytes = None

    with export1:
        st.download_button(
            "⬇️ Télécharger le classeur Excel complet",
            data=excel_bytes,
            file_name="analyse_photovoltaique_cma.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    with export2:
        csv_data = daily_export.to_csv(
            index=False,
            sep=";",
        ).encode("utf-8-sig")

        st.download_button(
            "⬇️ Télécharger les consommations journalières",
            data=csv_data,
            file_name="consommations_journalieres.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("---")
    st.subheader("Rapport pédagogique CMA")

    if pdf_report_bytes:
        pdf_filename_company = (
            company_name.strip().replace(" ", "_")
            if company_name.strip()
            else "entreprise"
        )

        st.download_button(
            "📄 Télécharger le rapport PDF CMA",
            data=pdf_report_bytes,
            file_name=(
                f"pre_diagnostic_photovoltaique_"
                f"{pdf_filename_company}.pdf"
            ),
            mime="application/pdf",
            use_container_width=True,
        )
    else:
        st.info(
            "Pour générer le rapport PDF, validez une adresse et "
            "assurez-vous que les données PVGIS sont disponibles."
        )

    st.markdown(
        """
        Le rapport PDF présente les résultats avec des explications simples,
        des graphiques et des recommandations adaptées à un échange avec
        l'artisan. Il reste un pré-diagnostic et ne remplace pas une étude
        technique ou économique complète.
        """
    )

    st.markdown(
        """
        Le classeur Excel contient :

        - la synthèse générale ;
        - les données traitées et normalisées à un pas de 1 heure ;
        - le profil horaire détaillé ;
        - les consommations journalières ;
        - les consommations mensuelles ;
        - le tableau moyen heure × jour de la semaine ;
        - la synthèse HP/HC hiver et été ;
        - le détail du classement tarifaire intervalle par intervalle ;
        - la synthèse financière du projet ;
        - la projection annuelle des flux financiers.
        """
    )
