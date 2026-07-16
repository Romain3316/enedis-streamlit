
import base64
from io import BytesIO
from pathlib import Path

import requests

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
    """Géocode une adresse avec le service officiel Géoplateforme / BAN."""
    address = address.strip()

    if len(address) < 5:
        raise ValueError("Veuillez saisir une adresse plus complète.")

    url = "https://data.geopf.fr/geocodage/search"
    params = {
        "text": address,
        "maximumResponses": 5,
    }

    response = requests.get(url, params=params, timeout=20)
    response.raise_for_status()
    payload = response.json()

    candidates = []

    # Format GeoJSON généralement renvoyé par le service search.
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
            or properties.get("fulltext")
            or properties.get("name")
            or address
        )

        score = properties.get("score")

        candidates.append(
            {
                "label": str(label),
                "latitude": latitude,
                "longitude": longitude,
                "score": score,
            }
        )

    # Prise en charge d'un éventuel autre format de réponse.
    for result in payload.get("results", []):
        longitude = result.get("x") or result.get("lon")
        latitude = result.get("y") or result.get("lat")

        if longitude is None or latitude is None:
            lonlat = result.get("lonlat")
            if isinstance(lonlat, str) and "," in lonlat:
                longitude, latitude = lonlat.split(",", 1)

        if longitude is None or latitude is None:
            continue

        candidates.append(
            {
                "label": (
                    result.get("fulltext")
                    or result.get("label")
                    or result.get("name")
                    or address
                ),
                "latitude": float(latitude),
                "longitude": float(longitude),
                "score": result.get("score"),
            }
        )

    # Déduplication simple.
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
            "Ajoutez le numéro, la rue, le code postal et la commune."
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
    Ajoute les heures astronomiques exactes de lever/coucher et la hauteur
    solaire à chaque relevé Enedis.
    """
    result = df.copy()
    location = Location(
        latitude=latitude,
        longitude=longitude,
        tz="Europe/Paris",
    )

    local_times = localize_paris(result["Horodate"])
    solar_position = location.get_solarposition(local_times)

    result["Hauteur_soleil_deg"] = solar_position[
        "apparent_elevation"
    ].to_numpy()

    unique_dates = pd.DatetimeIndex(
        pd.to_datetime(result["Horodate"].dt.normalize().unique())
    )
    local_midnights = unique_dates.tz_localize(
        "Europe/Paris",
        ambiguous=True,
        nonexistent="shift_forward",
    )

    sun_events = location.get_sun_rise_set_transit(
        local_midnights,
        method="spa",
    )

    sun_table = pd.DataFrame(
        {
            "Date_solaire": unique_dates,
            "Lever_soleil": (
                sun_events["sunrise"]
                .dt.tz_convert("Europe/Paris")
                .dt.tz_localize(None)
                .to_numpy()
            ),
            "Coucher_soleil": (
                sun_events["sunset"]
                .dt.tz_convert("Europe/Paris")
                .dt.tz_localize(None)
                .to_numpy()
            ),
            "Midi_solaire": (
                sun_events["transit"]
                .dt.tz_convert("Europe/Paris")
                .dt.tz_localize(None)
                .to_numpy()
            ),
        }
    )

    result["Date_solaire"] = result["Horodate"].dt.normalize()
    result = result.merge(
        sun_table,
        on="Date_solaire",
        how="left",
    )

    result["Soleil_leve"] = (
        (result["Horodate"] >= result["Lever_soleil"])
        & (result["Horodate"] <= result["Coucher_soleil"])
        & (result["Hauteur_soleil_deg"] > 0)
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
    st.markdown("## 3. Localisation solaire")

    company_address = st.text_input(
        "Adresse complète de l'entreprise",
        placeholder="Ex. 46 avenue du Général de Larminat, 33000 Bordeaux",
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
            f"Latitude : {selected_location['latitude']:.6f}\n\n"
            f"Longitude : {selected_location['longitude']:.6f}"
        )

    st.markdown("### Paramètres photovoltaïques")

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

    irradiation_threshold = st.slider(
        "Seuil d'irradiation significative",
        min_value=0,
        max_value=500,
        value=50,
        step=10,
        format="%d W/m²",
    )

    st.caption(
        "Le lever et le coucher sont calculés pour chaque date à partir "
        "des coordonnées exactes. L'irradiation et la production sont "
        "estimées à partir de PVGIS."
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

solar_analysis_available = selected_location is not None
pvgis_available = False
solar_error = None
solar_daily_df = pd.DataFrame()

daylight_kwh = np.nan
daylight_share = np.nan
irradiated_kwh = np.nan
irradiated_share = np.nan
pvgis_production_kwh = np.nan
self_consumed_kwh = np.nan
self_consumption_rate = np.nan
self_sufficiency_rate = np.nan

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

            irradiated_mask = (
                filtered_df["Irradiation_Wm2"]
                >= irradiation_threshold
            )

            irradiated_kwh = filtered_df.loc[
                irradiated_mask,
                "Energie_kWh",
            ].sum()

            irradiated_share = (
                irradiated_kwh / total_kwh * 100
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
    tab_solar,
    tab_profiles,
    tab_daily,
    tab_quality,
    tab_export,
) = st.tabs(
    [
        "📊 Tableau de bord",
        "☀️ Analyse solaire",
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
        "Part pendant le jour",
        (
            f"{format_fr(daylight_share, 1)} %"
            if solar_analysis_available
            else "Adresse requise"
        ),
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
            f"{selected_location['longitude']:.6f}**"
        )

        if solar_error:
            st.warning(solar_error)

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
        )
        s2.metric(
            "Coucher au début de période",
            first_date_row["Coucher_soleil"].strftime("%H:%M"),
        )
        s3.metric(
            "Lever en fin de période",
            last_date_row["Lever_soleil"].strftime("%H:%M"),
        )
        s4.metric(
            "Coucher en fin de période",
            last_date_row["Coucher_soleil"].strftime("%H:%M"),
        )

        k1, k2, k3, k4 = st.columns(4)

        k1.metric(
            "Consommation pendant le jour",
            f"{format_fr(daylight_kwh, 0)} kWh",
            f"{format_fr(daylight_share, 1)} %",
        )

        k2.metric(
            f"Conso. avec ≥ {irradiation_threshold} W/m²",
            (
                f"{format_fr(irradiated_kwh, 0)} kWh"
                if pvgis_available
                else "PVGIS indisponible"
            ),
            (
                f"{format_fr(irradiated_share, 1)} %"
                if pvgis_available
                else None
            ),
        )

        k3.metric(
            f"Production estimée {pv_peak_kwp:g} kWc",
            (
                f"{format_fr(pvgis_production_kwh, 0)} kWh"
                if pvgis_available
                else "PVGIS indisponible"
            ),
        )

        k4.metric(
            "Taux d'autoproduction estimé",
            (
                f"{format_fr(self_sufficiency_rate, 1)} %"
                if pvgis_available
                else "PVGIS indisponible"
            ),
        )

        if pvgis_available:
            a1, a2 = st.columns(2)

            with a1:
                st.metric(
                    "Énergie PV autoconsommée estimée",
                    f"{format_fr(self_consumed_kwh, 0)} kWh",
                )

            with a2:
                st.metric(
                    "Taux d'autoconsommation estimé",
                    f"{format_fr(self_consumption_rate, 1)} %",
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

            fig_irradiation.add_hline(
                y=irradiation_threshold,
                line_dash="dash",
                annotation_text=(
                    f"Seuil {irradiation_threshold} W/m²"
                ),
            )

            fig_irradiation.update_layout(
                template="plotly_white",
                hovermode="x unified",
            )

            st.plotly_chart(
                fig_irradiation,
                use_container_width=True,
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
                "Adresse de l'entreprise",
                "Latitude",
                "Longitude",
                "Part de consommation pendant le jour (%)",
                (
                    f"Part avec irradiation ≥ "
                    f"{irradiation_threshold} W/m² (%)"
                ),
                "Puissance photovoltaïque étudiée (kWc)",
                "Production PVGIS estimée (kWh)",
                "Taux d'autoconsommation estimé (%)",
                "Taux d'autoproduction estimé (%)",
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
                    irradiated_share
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
                    self_consumption_rate
                    if pvgis_available
                    else ""
                ),
                (
                    self_sufficiency_rate
                    if pvgis_available
                    else ""
                ),
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
