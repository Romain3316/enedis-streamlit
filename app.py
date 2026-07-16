
import base64
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


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
    f"""
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
                radial-gradient(circle at top right, rgba(229,57,53,.035), transparent 28rem),
                #FFFFFF;
            color: var(--cma-text);
        }

        .block-container {
            max-width: 1500px;
            padding-top: 1rem;
            padding-bottom: 3rem;
        }

        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg, #F8F9FB 0%, #EEF2F6 100%);
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
            min-height: 215px;
            padding: 38px 42px;
            overflow: hidden;
            color: white;
            border-radius: 22px;
            margin-bottom: 22px;
            background:
                linear-gradient(
                    116deg,
                    var(--cma-blue) 0%,
                    var(--cma-blue) 73%,
                    var(--cma-red) 73%,
                    var(--cma-red) 100%
                );
            box-shadow:
                0 20px 45px rgba(23,54,93,.18),
                0 4px 12px rgba(23,54,93,.09);
        }

        .cma-header::before {
            content: "☀";
            position: absolute;
            left: 58%;
            top: -52px;
            font-size: 210px;
            line-height: 1;
            color: rgba(255,255,255,.045);
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
            border: 3px solid rgba(255,255,255,.05);
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
            margin-bottom: 12px;
            padding: 6px 11px;
            border: 1px solid rgba(255,255,255,.28);
            border-radius: 999px;
            background: rgba(255,255,255,.10);
            font-size: .78rem;
            font-weight: 800;
            letter-spacing: .08em;
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
            text-shadow: 0 4px 16px rgba(0,0,0,.28);
        }

        .cma-title-line {
            width: 96px;
            height: 5px;
            margin: 20px 0 17px;
            background: #FFFFFF;
            border-radius: 999px;
            box-shadow: 0 2px 8px rgba(0,0,0,.12);
        }

        .cma-header p {
            margin: 0;
            max-width: 680px;
            color: rgba(255,255,255,.96);
            font-size: 1.2rem;
            font-weight: 500;
            line-height: 1.5;
        }

        .cma-logo {
            position: relative;
            z-index: 3;
            display: flex;
            align-items: center;
            justify-content: center;
            flex: 0 0 320px;
            width: 320px;
            min-height: 142px;
            padding: 18px 22px;
            background: #FFFFFF;
            color: var(--cma-blue);
            border: 1px solid rgba(255,255,255,.65);
            border-radius: 19px;
            text-align: center;
            font-size: 1.35rem;
            font-weight: 900;
            box-shadow:
                0 15px 32px rgba(0,0,0,.16),
                inset 0 0 0 1px rgba(23,54,93,.04);
        }

        .cma-logo img {
            display: block;
            width: 100%;
            max-width: 280px;
            max-height: 112px;
            object-fit: contain;
        }

        /* Carte d'introduction */
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
            box-shadow: 0 8px 24px rgba(23,54,93,.07);
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

        /* Cartes de parcours */
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 18px;
            margin: 8px 0 24px;
        }

        .feature-card {
            position: relative;
            min-height: 150px;
            padding: 24px;
            overflow: hidden;
            background: #FFFFFF;
            border: 1px solid var(--cma-border);
            border-radius: 18px;
            box-shadow: 0 8px 22px rgba(23,54,93,.07);
            transition:
                transform .22s ease,
                box-shadow .22s ease,
                border-color .22s ease;
        }

        .feature-card:hover {
            transform: translateY(-4px);
            border-color: rgba(23,54,93,.25);
            box-shadow: 0 16px 32px rgba(23,54,93,.12);
        }

        .feature-card::after {
            content: "";
            position: absolute;
            right: -28px;
            bottom: -35px;
            width: 110px;
            height: 110px;
            border-radius: 50%;
            background: rgba(23,54,93,.035);
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

        /* KPI */
        div[data-testid="stMetric"] {
            min-height: 122px;
            padding: 18px 19px;
            background: #FFFFFF;
            border: 1px solid var(--cma-border);
            border-top: 4px solid var(--cma-blue);
            border-radius: 15px;
            box-shadow: 0 7px 20px rgba(23,54,93,.06);
        }

        div[data-testid="stMetricLabel"] {
            color: #697589;
            font-size: .82rem;
            font-weight: 750;
            text-transform: uppercase;
            letter-spacing: .03em;
        }

        div[data-testid="stMetricValue"] {
            color: var(--cma-blue);
            font-size: 1.85rem;
            font-weight: 850;
        }

        /* Onglets */
        button[data-baseweb="tab"] {
            font-weight: 750;
        }

        /* Boutons */
        .stButton > button {
            min-height: 44px;
            padding: .65rem 1.15rem;
            background: var(--cma-red);
            color: white;
            border: 0;
            border-radius: 11px;
            font-weight: 750;
            box-shadow: 0 5px 13px rgba(229,57,53,.18);
        }

        .stButton > button:hover {
            background: var(--cma-red-dark);
            color: white;
            transform: translateY(-1px);
        }

        .stDownloadButton > button {
            min-height: 46px;
            background: var(--cma-blue);
            color: white;
            border: 0;
            border-radius: 11px;
            font-weight: 750;
            box-shadow: 0 5px 13px rgba(23,54,93,.16);
        }

        .stDownloadButton > button:hover {
            background: var(--cma-blue-dark);
            color: white;
            transform: translateY(-1px);
        }

        h2, h3 {
            color: var(--cma-blue);
        }

        @media (max-width: 900px) {
            .cma-header {
                flex-direction: column;
                align-items: stretch;
                min-height: auto;
                padding: 30px 26px;
                background:
                    linear-gradient(
                        165deg,
                        var(--cma-blue) 0%,
                        var(--cma-blue) 70%,
                        var(--cma-red) 70%,
                        var(--cma-red) 100%
                    );
            }

            .cma-logo {
                width: 100%;
                max-width: 360px;
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
# FONCTIONS
# ============================================================

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
            f'<div class="cma-logo">'
            f'<img src="{logo_uri}" alt="Logo CMA Nouvelle-Aquitaine">'
            f"</div>"
        )
    else:
        logo_html = (
            '<div class="cma-logo">'
            'CMA<br><span style="font-size:.76rem;font-weight:650;">'
            "NOUVELLE-AQUITAINE</span></div>"
        )

    st.markdown(
        f"""
        <div class="cma-header">
            <div class="cma-header-content">
                <div class="cma-eyebrow">CMA Nouvelle-Aquitaine · Outil métier</div>
                <h1>Pré-diagnostic photovoltaïque</h1>
                <div class="cma-title-line"></div>
                <p>Analyse automatisée des courbes de charge Enedis et génération
                d'indicateurs d'aide au diagnostic énergétique.</p>
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
            <strong>Analysez votre fichier Enedis en quelques secondes.</strong><br>
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
                <p>Déposez un export Enedis au format CSV ou Excel depuis le panneau latéral.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon">📊</div>
                <h3>2. Analyser</h3>
                <p>Obtenez automatiquement les profils de charge, consommations et contrôles qualité.</p>
            </div>
            <div class="feature-card">
                <div class="feature-icon">📥</div>
                <h3>3. Exporter</h3>
                <p>Téléchargez un classeur Excel structuré, prêt à être exploité dans le diagnostic.</p>
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
    st.markdown("## 3. Plage solaire")

    solar_start = st.slider(
        "Début",
        min_value=0,
        max_value=23,
        value=8,
        format="%dh",
    )

    solar_end = st.slider(
        "Fin",
        min_value=1,
        max_value=24,
        value=18,
        format="%dh",
    )

    st.caption(
        "Cette plage sert uniquement à mesurer la part de consommation "
        "située pendant des heures potentiellement favorables au solaire."
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

total_kwh = filtered_df["Energie_kWh"].sum()
average_daily_kwh = daily_df["Consommation_kWh"].mean()
median_daily_kwh = daily_df["Consommation_kWh"].median()
maximum_power_kw = filtered_df["Puissance_kW"].max()
mean_power_kw = filtered_df["Puissance_kW"].mean()

solar_mask = (
    (filtered_df["Horodate"].dt.hour >= solar_start)
    & (filtered_df["Horodate"].dt.hour < solar_end)
)

solar_kwh = filtered_df.loc[
    solar_mask,
    "Energie_kWh",
].sum()

solar_share = (
    solar_kwh / total_kwh * 100
    if total_kwh
    else 0
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
    tab_profiles,
    tab_daily,
    tab_quality,
    tab_export,
) = st.tabs(
    [
        "📊 Tableau de bord",
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
    metric1, metric2, metric3, metric4, metric5 = st.columns(5)

    metric1.metric(
        "Consommation totale",
        f"{format_fr(total_kwh, 0)} kWh",
    )

    metric2.metric(
        "Moyenne journalière",
        f"{format_fr(average_daily_kwh, 1)} kWh",
    )

    metric3.metric(
        "Pic de puissance",
        f"{format_fr(maximum_power_kw, 1)} kW",
    )

    metric4.metric(
        f"Part {solar_start}h–{solar_end}h",
        f"{format_fr(solar_share, 1)} %",
    )

    metric5.metric(
        "Facteur de charge",
        f"{format_fr(load_factor, 1)} %",
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
                        f"Entre {solar_start}h et {solar_end}h",
                        "Hors plage solaire",
                    ],
                    values=[
                        solar_kwh,
                        max(total_kwh - solar_kwh, 0),
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
                f"Part entre {solar_start}h et {solar_end}h (%)",
                "Facteur de charge (%)",
                "Horodatages en doublon",
                "Jours atypiques",
            ],
            "Valeur": [
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
                solar_share,
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
    hourly_standardized_export = hourly_standardized_export[
        ["Unité", "Horodate", "Valeur", "Pas"]
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

    excel_bytes = make_excel_export(
        hourly_standardized_export,
        hourly_export,
        daily_export,
        monthly_export,
        weekday_hour_matrix.round(3),
        summary_df,
    )

    export1, export2 = st.columns(2)

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

    st.markdown(
        """
        Le classeur Excel contient :

        - la synthèse générale ;
        - les données traitées et normalisées à un pas de 1 heure ;
        - le profil horaire détaillé ;
        - les consommations journalières ;
        - les consommations mensuelles ;
        - le tableau moyen heure × jour de la semaine.
        """
    )
