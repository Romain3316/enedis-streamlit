"""Compact CMA workspace styling, layered after Streamlit's base theme."""

import base64
from html import escape
from pathlib import Path
import streamlit as st


def apply_theme():
    st.markdown('''<style>
    .stApp {background:#F6F8FB !important;color:#17365D !important;}
    .block-container {padding-top:1.4rem !important; max-width:1680px !important; padding-bottom:2rem;}
    [data-testid="stSidebar"] {background:#F0F4F8 !important;border-right:1px solid #DCE4ED;}
    [data-testid="stSidebar"] > div:first-child {padding-top:1.2rem;}
    h1,h2,h3 {color:#17365D !important;letter-spacing:-.025em !important;}
    h1 {font-size:1.7rem !important;} h2 {font-size:1.35rem !important;} h3 {font-size:1.05rem !important;}
    [data-testid="stCaptionContainer"] {color:#64748B !important;}
    .cma-topbar {display:flex;gap:18px;align-items:center;min-height:64px;margin-bottom:10px;}
    .cma-topbar img {width:120px;max-height:56px;object-fit:contain;}
    .cma-topbar-title {font-size:1.25rem;font-weight:750;color:#17365D;line-height:1.3;}
    .cma-topbar-subtitle {font-size:.82rem;color:#64748B;margin-top:4px;}
    .cma-dossier {padding:14px 0;border-top:1px solid #DCE4ED;border-bottom:1px solid #DCE4ED;margin:16px 0;}
    .cma-dossier strong {display:block;color:#17365D;font-size:1rem;}
    .cma-dossier p {font-size:.8rem;color:#64748B;margin:6px 0;}
    .cma-status {display:inline-block;border-radius:4px;padding:3px 8px;font-size:.75rem;background:#E1E8F0;color:#17365D;}
    div[data-testid="stMetric"] {background:white !important;min-height:96px !important;padding:14px 16px !important;
      border:1px solid #DCE4ED !important;border-radius:7px !important;box-shadow:none !important;}
    [data-testid="stMetricValue"] {font-size:1.65rem !important;color:#17365D !important;}
    [data-testid="stMetricLabel"] {text-transform:none !important;font-size:.85rem !important;font-weight:500 !important;}
    [data-testid="stVerticalBlockBorderWrapper"] {border-radius:7px !important;box-shadow:none !important;}
    [data-baseweb="tab-list"] {gap:18px !important;background:transparent !important;border-bottom:1px solid #DCE4ED;padding:0 !important;}
    button[data-baseweb="tab"] {font-size:.9rem !important;border-radius:0 !important;background:transparent !important;padding:10px 0 !important;color:#53667F !important;}
    button[data-baseweb="tab"][aria-selected="true"] {color:#17365D !important;font-weight:700 !important;background:transparent !important;}
    [data-baseweb="tab-highlight"] {background:#E53935 !important;height:3px !important;}
    .stButton button,.stDownloadButton button {border-radius:6px !important;box-shadow:none !important;min-height:40px;}
    button[kind="primary"] {background:#17365D !important;border-color:#17365D !important;color:white !important;}
    [data-testid="stDataFrame"] {border-radius:5px !important;box-shadow:none !important;}
    .cma-heat-table {width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed;}
    .cma-heat-table th,.cma-heat-table td {padding:3px 4px !important;line-height:17px;height:23px;border:1px solid white;text-align:center;}
    .cma-heat-table thead th {background:#17365D;color:white;font-weight:600;}
    .cma-heat-table tbody th {background:#EDF2F7;color:#17365D;font-weight:500;}
    .intro-card,.feature-card,.score-card,.pedagogy-card {box-shadow:none !important;border-radius:7px !important;}
    .feature-card:hover {transform:none !important;}
    .heat-legend {display:flex;align-items:center;gap:10px;font-size:.75rem;color:#53667F;margin:6px 0 16px;}
    .heat-legend span {width:180px;height:10px;background:linear-gradient(90deg,#63BE7B,#A9D26D 30%,#FFEB84 50%,#F6B26B 72%,#F8696B);border-radius:2px;}
    @media(max-width:800px) {.block-container {padding:1rem !important;} .cma-topbar img{width:90px;} .cma-topbar-title{font-size:1rem;}}
    </style>''', unsafe_allow_html=True)


def compact_header():
    path = Path(__file__).parent / 'logo_cma.png'
    logo = f'<img src="data:image/png;base64,{base64.b64encode(path.read_bytes()).decode()}" alt="CMA Nouvelle-Aquitaine">' if path.exists() else '<strong>CMA</strong>'
    st.markdown(f'<div class="cma-topbar">{logo}<div><div class="cma-topbar-title">Analyse énergétique</div><div class="cma-topbar-subtitle">CMA Nouvelle-Aquitaine · Outil métier</div></div></div>',unsafe_allow_html=True)


def dossier_summary():
    name = escape(st.session_state.get('company_name') or 'Nouveau dossier')
    status = escape(st.session_state.get('report_status','Brouillon'))
    st.markdown(f'<div class="cma-dossier"><p>VOTRE DOSSIER</p><strong>{name}</strong><p><span class="cma-status">{status}</span></p></div>',unsafe_allow_html=True)


def heat_legend():
    st.markdown('<div class="heat-legend">Puissance faible <span></span> Puissance élevée</div>',unsafe_allow_html=True)
    st.caption("Échelle relative à la période sélectionnée. Une valeur élevée ne signifie pas nécessairement une anomalie.")


def open_report():
    st.session_state['_workspace_page'] = 'Rapport'


def render_profile_matrix(matrix):
    from heatmap_colors import style_matrix
    display = matrix.rename(index=lambda h: f"{int(h):02d}h")
    st.markdown(style_matrix(display).set_table_attributes('class="cma-heat-table"').to_html(), unsafe_allow_html=True)
    heat_legend()
