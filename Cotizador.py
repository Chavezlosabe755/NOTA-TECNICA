"""
Cotizador - Seguro de Casa Habitación
Bouclier Seguros de Daños
======================================
Requiere: RESULTADOS_NOTA_TECNICA.xlsx en el mismo directorio
Instalar: pip install streamlit pandas openpyxl
Correr:   streamlit run cotizador.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import date
import random
import io
try:
    from fpdf import FPDF
    FPDF_OK = True
except ImportError:
    FPDF_OK = False

st.set_page_config(
    page_title="Cotizador | Casa Habitación · Bouclier",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS — Branding Bouclier
# Color principal: guinda #7B1C35
# Acento beige:   #C4B5A5
# Fondo claro:    #F8F5F2
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #F8F5F2;
    color: #1a1a2e;
}

.stApp { background-color: #F8F5F2; }

/* Sidebar — oscuro */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div,
[data-testid="stSidebar"] > div > div,
section[data-testid="stSidebar"] > div {
    background-color: #7B1C35 !important;
    border-right: none !important;
}
[data-testid="stSidebar"] * { color: #f5ede8 !important; }
[data-testid="stSidebar"] input { color: #1a1a2e !important; }
[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] span,
[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] div,
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] { color: #1a1a2e !important; }
[data-testid="stSidebar"] .stRadio span { color: #f5ede8 !important; }
[data-testid="stSidebar"] .stCheckbox span { color: #f5ede8 !important; }
[data-testid="stSidebar"] .stNumberInput input { color: #1a1a2e !important; }
[data-testid="stSidebar"] .seccion-titulo {
    color: #C4B5A5 !important;
    border-bottom-color: rgba(196,181,165,0.4) !important;
}
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] select,
[data-testid="stSidebar"] .stSelectbox > div,
[data-testid="stSidebar"] .stNumberInput > div > div {
    background-color: #ffffff !important;
    border-color: #C4B5A5 !important;
    color: #1a1a2e !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] .stButton > button {
    background-color: #7B1C35 !important;
    color: #f5ede8 !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 1rem !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.03em;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background-color: #9B2C45 !important;
}

/* Botones +/- de NumberInput en sidebar */
[data-testid="stSidebar"] button[data-testid="stNumberInputStepDown"],
[data-testid="stSidebar"] button[data-testid="stNumberInputStepUp"],
[data-testid="stSidebar"] .stNumberInput button {
    background-color: #C4B5A5 !important;
    color: #1a1a2e !important;
    border: none !important;
    border-radius: 4px !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] button[data-testid="stNumberInputStepDown"]:hover,
[data-testid="stSidebar"] button[data-testid="stNumberInputStepUp"]:hover,
[data-testid="stSidebar"] .stNumberInput button:hover {
    background-color: #7B1C35 !important;
    color: #f5ede8 !important;
}
/* Botón Calcular Cotización (primary) */
[data-testid="stSidebar"] .stButton > button[kind="primary"],
[data-testid="stSidebar"] .stButton > button[data-testid="baseButton-primary"] {
    background-color: #C4B5A5 !important;
    color: #1a1a2e !important;
    border: 2px solid #C4B5A5 !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
    background-color: #f5ede8 !important;
    color: #7B1C35 !important;
}

[data-testid="stSidebar"] label { color: #f5ede8 !important; font-size: 0.82rem !important; }
[data-testid="stSidebar"] .stRadio label { font-size: 0.88rem !important; }
[data-testid="stSidebar"] .stCheckbox label { font-size: 0.85rem !important; }

/* Header strip */
.header-strip {
    background: linear-gradient(135deg, #7B1C35 60%, #5a1326 100%);
    border-radius: 16px;
    padding: 0;
    margin-bottom: 1.5rem;
    overflow: hidden;
    position: relative;
    display: flex;
    align-items: stretch;
    min-height: 110px;
}
.header-deco {
    width: 120px;
    min-height: 110px;
    background: linear-gradient(160deg, #C4B5A5 0%, #a89888 100%);
    clip-path: polygon(0 0, 100% 0, 60% 100%, 0 100%);
    flex-shrink: 0;
}
.header-content {
    flex: 1;
    padding: 1.4rem 1.8rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.header-text h1 {
    font-family: 'Cormorant Garamond', serif;
    font-size: 2rem;
    font-weight: 700;
    color: #f5ede8;
    margin: 0 0 0.2rem 0;
    letter-spacing: 0.02em;
}
.header-text p {
    font-size: 0.78rem;
    color: #C4B5A5;
    margin: 0;
    letter-spacing: 0.12em;
    text-transform: uppercase;
}

/* Logo SVG Bouclier */
.logo-wrap {
    width: 90px; height: 90px;
    background: #7B1C35;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    flex-shrink: 0;
    overflow: hidden;
    border: 3px solid #C4B5A5;
}
.logo-wrap img { width: 80px; height: 80px; object-fit: contain; filter: invert(1); }

/* Sección título */
.seccion-titulo {
    font-size: 0.68rem;
    color: #7B1C35;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    font-weight: 600;
    margin: 1.4rem 0 0.6rem 0;
    border-bottom: 2px solid #7B1C35;
    padding-bottom: 0.35rem;
}

/* Cards */
.card {
    background: #fff;
    border: 1px solid #e8ddd8;
    border-radius: 12px;
    padding: 1.3rem 1.6rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 8px rgba(123,28,53,0.06);
}
.card-label {
    font-size: 0.72rem; color: #9a8a84;
    text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 0.3rem;
}
.card-value {
    font-family: 'Cormorant Garamond', serif;
    font-size: 2.2rem; color: #7B1C35; line-height: 1.1; font-weight: 700;
}
.card-sub { font-size: 0.8rem; color: #9a8a84; margin-top: 0.25rem; }

/* Métricas de datos */
[data-testid="stMetric"] {
    background: #fff;
    border: 1px solid #e8ddd8;
    border-radius: 10px;
    padding: 0.8rem 1rem !important;
    box-shadow: 0 1px 4px rgba(123,28,53,0.05);
}
[data-testid="stMetricLabel"] { color: #9a8a84 !important; font-size: 0.72rem !important; text-transform: uppercase; letter-spacing: 0.08em; }
[data-testid="stMetricValue"] { color: #1a1a2e !important; font-size: 1.1rem !important; font-weight: 600 !important; }

/* Tabla de cotización */
.tbl { width:100%; border-collapse:collapse; font-size:0.83rem; }
.tbl th {
    background: #7B1C35;
    color: #f5ede8;
    font-size: 0.68rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    padding: 0.65rem 0.8rem;
    text-align: right;
    font-weight: 500;
}
.tbl th:first-child, .tbl th:nth-child(2), .tbl th:nth-child(3) { text-align: left; }
.tbl th.col-prima { background: #5a1326 !important; color: #f5c8b8 !important; font-weight: 700 !important; }
.tbl td {
    padding: 0.55rem 0.8rem;
    border-bottom: 1px solid #f0e8e3;
    color: #2a1a1e;
    text-align: right;
    vertical-align: middle;
}
.tbl td:first-child, .tbl td:nth-child(2), .tbl td:nth-child(3) { text-align: left; }
.tbl td.col-prima { background: #fdf0ee !important; color: #7B1C35 !important; font-weight: 700 !important; }
.tbl tr:hover td { background: #fdf5f2; }
.tbl tr:hover td.col-prima { background: #f9e4e0 !important; }
.tbl tr.total-row td {
    background: #fdf5f2;
    color: #7B1C35;
    font-weight: 700;
    border-top: 2px solid #7B1C35;
    border-bottom: none;
    font-size: 0.88rem;
}
.tbl tr.total-row td.col-prima { background: #7B1C35 !important; color: #f5ede8 !important; font-size: 0.95rem !important; }
.badge-si {
    background: #e8f5ee; color: #1a7a45;
    border-radius: 4px; padding: 2px 8px;
    font-size: 0.72rem; font-weight: 700;
}
.badge-no {
    background: #fdecea; color: #c0392b;
    border-radius: 4px; padding: 2px 8px;
    font-size: 0.72rem; font-weight: 700;
}

/* Desglose */
.desglose-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 0.45rem 0; border-bottom: 1px solid #f0e8e3; font-size: 0.88rem;
}
.desglose-row:last-child { border-bottom: none; }
.desglose-label { color: #7a6a6e; }
.desglose-valor { color: #1a1a2e; font-weight: 500; }
.desglose-total {
    display: flex; justify-content: space-between; align-items: center;
    padding: 0.6rem 0; font-weight: 700; color: #7B1C35;
    border-top: 2px solid #7B1C35; margin-top: 0.3rem; font-size: 0.95rem;
}

/* Alerta */
.alerta {
    background: #fdf5f2;
    border-left: 3px solid #7B1C35;
    border-radius: 6px;
    padding: 0.8rem 1rem;
    font-size: 0.82rem;
    color: #7a6a6e;
    margin-top: 1rem;
}

/* Sidebar logo */
.sidebar-logo-wrap {
    display: flex; align-items: center; gap: 0.8rem;
    padding: 0.5rem 0 1rem 0; margin-bottom: 0.5rem;
}
.sidebar-logo-wrap .sb-logo {
    width: 52px; height: 52px; border-radius: 50%;
    background: transparent;
    display: flex; align-items: center; justify-content: center;
    border: 2px solid rgba(196,181,165,0.5); flex-shrink: 0;
    overflow: hidden;
}
.sidebar-logo-wrap .sb-logo img { width: 44px; height: 44px; object-fit: contain; filter: brightness(0) invert(1); }
.sidebar-brand { font-family: 'Cormorant Garamond', serif; font-size: 1.1rem; font-weight: 700; color: #C4B5A5 !important; }
.sidebar-sub { font-size: 0.68rem; color: #C4B5A5 !important; letter-spacing: 0.1em; text-transform: uppercase; }


/* Sidebar negro — mejor visibilidad */
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stDateInput input {
    background-color: #ffffff !important;
    border-color: #C4B5A5 !important;
    color: #1a1a2e !important;
}
[data-testid="stSidebar"] .stSelectbox > div > div,
[data-testid="stSidebar"] .stNumberInput input {
    background-color: #ffffff !important;
    border-color: #C4B5A5 !important;
    color: #1a1a2e !important;
}
[data-testid="stSidebar"] .stRadio > div { gap: 0.3rem; }

footer { visibility: hidden; }
#MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# SVG del logo Bouclier (águila/fénix estilizado con círculo y texto)
LOGO_IMG = """
<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" style="width:80px;height:80px">
  <circle cx="100" cy="100" r="96" stroke="#C4B5A5" stroke-width="2" fill="none"/>
  <circle cx="100" cy="100" r="84" stroke="#C4B5A5" stroke-width="1.2" fill="none"/>
  <path d="M28 100 A72 72 0 0 1 36 65" stroke="#f5ede8" stroke-width="3.5" stroke-linecap="round" fill="none"/>
  <path d="M164 65 A72 72 0 0 1 172 100" stroke="#f5ede8" stroke-width="3.5" stroke-linecap="round" fill="none"/>
  <path d="M30 118 A72 72 0 0 0 42 148" stroke="#f5ede8" stroke-width="3.5" stroke-linecap="round" fill="none"/>
  <path d="M158 148 A72 72 0 0 0 170 118" stroke="#f5ede8" stroke-width="3.5" stroke-linecap="round" fill="none"/>
  <circle cx="25" cy="100" r="3.5" fill="#f5ede8"/>
  <circle cx="175" cy="100" r="3.5" fill="#f5ede8"/>
  <path id="arc1" d="M 22 88 A 80 80 0 0 1 178 88" fill="none"/>
  <text font-family="Georgia,serif" font-size="15" font-weight="bold" fill="#f5ede8" letter-spacing="5">
    <textPath href="#arc1" startOffset="12%">BOUCLIER</textPath>
  </text>
  <path d="M80 75 Q76 66 84 60 Q92 54 104 58 Q116 54 122 62 Q130 70 124 80 Q120 87 112 86" stroke="#f5ede8" stroke-width="2.2" stroke-linejoin="round" fill="none"/>
  <path d="M80 75 Q76 84 82 90 Q90 96 104 96 Q108 104 100 108 Q90 110 82 102 Q72 92 74 78" stroke="#f5ede8" stroke-width="2.2" stroke-linejoin="round" fill="none"/>
  <circle cx="102" cy="64" r="2.5" fill="#f5ede8"/>
  <path d="M108 58 Q116 46 126 44 Q120 54 128 62" stroke="#f5ede8" stroke-width="2.2" stroke-linecap="round" fill="none"/>
  <path d="M112 56 Q124 38 138 36 Q130 50 136 62" stroke="#f5ede8" stroke-width="2" stroke-linecap="round" fill="none"/>
  <path d="M122 80 Q136 72 148 62 Q142 76 152 80 Q142 88 150 96 Q138 94 142 106 Q130 102 128 112" stroke="#f5ede8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <path d="M82 90 Q78 102 76 116 Q74 128 80 140" stroke="#f5ede8" stroke-width="2.5" stroke-linecap="round" fill="none"/>
  <path d="M104 96 Q108 108 106 122 Q104 134 100 144" stroke="#f5ede8" stroke-width="2" stroke-linecap="round" fill="none"/>
  <path d="M80 140 Q74 150 68 162 Q78 156 74 168 Q86 158 82 170 Q92 160 90 172 Q100 160 100 144" stroke="#f5ede8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <path d="M100 144 Q110 160 118 172 Q108 160 112 168 Q104 158 106 166 Q96 156 100 144" stroke="#f5ede8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <path d="M100 140 Q90 130 86 118 Q84 130 80 140 Q90 136 100 140 Q110 136 120 140 Q116 130 114 118 Q110 130 100 140Z" stroke="#f5ede8" stroke-width="1.8" stroke-linejoin="round" fill="none"/>
</svg>
"""

LOGO_IMG_SMALL = """
<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" style="width:36px;height:36px">
  <circle cx="100" cy="100" r="94" stroke="#C4B5A5" stroke-width="3" fill="none"/>
  <path id="arc2" d="M 22 88 A 80 80 0 0 1 178 88" fill="none"/>
  <text font-family="Georgia,serif" font-size="15" font-weight="bold" fill="#f5ede8" letter-spacing="5">
    <textPath href="#arc2" startOffset="12%">BOUCLIER</textPath>
  </text>
  <path d="M80 75 Q76 66 84 60 Q92 54 104 58 Q116 54 122 62 Q130 70 124 80 Q120 87 112 86" stroke="#f5ede8" stroke-width="2.5" stroke-linejoin="round" fill="none"/>
  <path d="M80 75 Q76 84 82 90 Q90 96 104 96 Q108 104 100 108 Q90 110 82 102 Q72 92 74 78" stroke="#f5ede8" stroke-width="2.5" stroke-linejoin="round" fill="none"/>
  <circle cx="102" cy="64" r="3" fill="#f5ede8"/>
  <path d="M108 58 Q116 46 126 44 Q120 54 128 62" stroke="#f5ede8" stroke-width="2.5" stroke-linecap="round" fill="none"/>
  <path d="M112 56 Q124 38 138 36 Q130 50 136 62" stroke="#f5ede8" stroke-width="2.2" stroke-linecap="round" fill="none"/>
  <path d="M122 80 Q136 72 148 62 Q142 76 152 80 Q142 88 150 96 Q138 94 142 106 Q130 102 128 112" stroke="#f5ede8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <path d="M82 90 Q78 102 76 116 Q74 128 80 140" stroke="#f5ede8" stroke-width="2.5" stroke-linecap="round" fill="none"/>
  <path d="M104 96 Q108 108 106 122 Q104 134 100 144" stroke="#f5ede8" stroke-width="2.2" stroke-linecap="round" fill="none"/>
  <path d="M80 140 Q74 150 68 162 Q78 156 74 168 Q86 158 82 170 Q92 160 90 172 Q100 160 100 144" stroke="#f5ede8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <path d="M100 144 Q110 160 118 172 Q108 160 112 168 Q104 158 106 166 Q96 156 100 144" stroke="#f5ede8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
</svg>
"""


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
EXCEL_PATH = "RESULTADOS_NOTA_TECNICA.xlsx"

@st.cache_data(show_spinner="Cargando tablas actuariales…")
def cargar_datos(path):
    xls = pd.ExcelFile(path)
    return (
        pd.read_excel(xls, "CR_Global"),
        pd.read_excel(xls, "Factor_Geografico"),
        pd.read_excel(xls, "Factor_Deducible"),
    )

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt_m(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"${v:,.2f}"

def factor_financiero(j, m):
    return m / sum((1 + j/m)**(-k) for k in range(m))

def get_cr_global(cr_df, cobertura, tipo_bien):
    mask = (cr_df["COBERTURA"] == cobertura) & (cr_df["TIPO DE BIEN"] == tipo_bien)
    f = cr_df[mask]
    if f.empty: return None
    v = f["CR_GLOBAL"].values[0]
    return None if pd.isna(v) or v == 0 else v

def get_f_geo(geo_df, cobertura, tipo_bien, estado):
    mask = (geo_df["COBERTURA"] == cobertura) & (geo_df["TIPO DE BIEN"] == tipo_bien) & (geo_df["ENTIDAD"] == estado)
    f = geo_df[mask]
    if f.empty: return 1.0
    v = f["FACTOR_GEO"].values[0]
    return 1.0 if pd.isna(v) or v <= 0 else v

def get_f_ded(ded_df, deducible):
    f = ded_df[ded_df["DEDUCIBLE"] == deducible]
    if f.empty: return 1.0
    v = f["f"].values[0]
    return 1.0 if pd.isna(v) else v


# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────
if not os.path.exists(EXCEL_PATH):
    st.error(f"⚠️ No se encontró **{EXCEL_PATH}**. Corre primero el notebook.")
    st.stop()
try:
    cr_global_df, factor_geo_df, factor_ded_df = cargar_datos(EXCEL_PATH)
except Exception as e:
    st.error(f"Error al leer el Excel: {e}"); st.stop()

coberturas_disp = sorted(cr_global_df["COBERTURA"].dropna().unique().tolist())
tipos_disp      = sorted(cr_global_df["TIPO DE BIEN"].dropna().unique().tolist())
estados_disp    = sorted(factor_geo_df["ENTIDAD"].dropna().unique().tolist())
deducibles_disp = sorted(factor_ded_df["DEDUCIBLE"].dropna().unique().tolist())
TIPOS_ORDEN     = {t: i+1 for i, t in enumerate(tipos_disp)}


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    # Logo + marca en sidebar
    st.markdown(f"""
    <div class="sidebar-logo-wrap">
        <div class="sb-logo">{LOGO_IMG_SMALL}</div>
        <div>
            <div class="sidebar-brand">Bouclier</div>
            <div class="sidebar-sub">Seguros de Daños</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Datos de póliza — fijos, no editables
    no_poliza     = f"INC{random.randint(1000000000, 9999999999)}"
    fecha_emision = date.today()
    fecha_inicio  = date.today()
    fecha_fin     = date.today().replace(year=date.today().year + 1)

    st.markdown('<div class="seccion-titulo">Datos de la Póliza</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div style="font-size:0.82rem;line-height:1.8;margin-bottom:0.5rem">
        <div style="display:flex;justify-content:space-between;padding:0.25rem 0;border-bottom:1px solid rgba(196,181,165,0.2)">
            <span style="color:#C4B5A5">No. Póliza</span>
            <span style="font-weight:600">{no_poliza}</span>
        </div>
        <div style="display:flex;justify-content:space-between;padding:0.25rem 0;border-bottom:1px solid rgba(196,181,165,0.2)">
            <span style="color:#C4B5A5">Fecha Emisión</span>
            <span style="font-weight:600">{fecha_emision.strftime('%d/%m/%Y')}</span>
        </div>
        <div style="display:flex;justify-content:space-between;padding:0.25rem 0;border-bottom:1px solid rgba(196,181,165,0.2)">
            <span style="color:#C4B5A5">Inicio Vigencia</span>
            <span style="font-weight:600">{fecha_inicio.strftime('%d/%m/%Y')}</span>
        </div>
        <div style="display:flex;justify-content:space-between;padding:0.25rem 0">
            <span style="color:#C4B5A5">Fin Vigencia</span>
            <span style="font-weight:600">{fecha_fin.strftime('%d/%m/%Y')}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="seccion-titulo">Datos del Asegurado</div>', unsafe_allow_html=True)
    nombre       = st.text_input("Nombre(s)")
    ap_paterno   = st.text_input("Apellido Paterno")
    ap_materno   = st.text_input("Apellido Materno")
    correo       = st.text_input("Correo electrónico")
    telefono     = st.text_input("Teléfono celular")

    st.markdown('<div class="seccion-titulo">Datos del Riesgo</div>', unsafe_allow_html=True)
    estado    = st.selectbox("Entidad Federativa", estados_disp)
    tipo_bien = st.selectbox("Tipo de Bien", tipos_disp)
    deducible = st.selectbox("Deducible ($)", deducibles_disp,
                              format_func=lambda x: f"${x:,.0f}")

    st.markdown('<div class="seccion-titulo">Forma de Pago</div>', unsafe_allow_html=True)
    forma_pago = st.radio("Periodicidad",
                          ["Anual", "Semestral", "Trimestral", "Mensual"])

    GASTO_ADMIN = 0.20; GASTO_ADQ = 0.05; MARGEN_UT = 0.06
    TASA_ANUAL  = 0.08; CARGOS    = GASTO_ADMIN + GASTO_ADQ + MARGEN_UT

    st.markdown('<div class="seccion-titulo">Coberturas a Cotizar</div>', unsafe_allow_html=True)
    st.caption("Activa cada cobertura e ingresa su Suma Asegurada")

    coberturas_config = []
    for i, cob in enumerate(coberturas_disp):
        col_chk, col_sa = st.columns([1, 2])
        with col_chk:
            activa = st.checkbox(cob, value=(i == 0), key=f"chk_{i}")
        with col_sa:
            sa = st.number_input(
                "SA", min_value=0.0, max_value=500_000_000.0,
                value=1_000_000.0 if i == 0 else 0.0,
                step=10_000.0, format="%.0f",
                key=f"sa_{i}", label_visibility="collapsed",
                disabled=not activa
            )
            if activa and sa > 0:
                st.caption(f"$ {sa:,.0f}")
        if activa and sa > 0:
            coberturas_config.append({"cobertura": cob, "sa": sa})

    cotizar = st.button("Calcular Cotización", use_container_width=True, type="primary")


# ─────────────────────────────────────────────
# HEADER PRINCIPAL
# ─────────────────────────────────────────────
st.markdown(f"""
<div class="header-strip">
    <div class="header-deco"></div>
    <div class="header-content">
        <div class="header-text">
            <h1>Cotizador · Casa Habitación</h1>
            <p>Ramo Incendio · Seguros de Daños Bouclier · México</p>
        </div>
        <div class="logo-wrap">{LOGO_IMG}</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# ÁREA PRINCIPAL
# ─────────────────────────────────────────────
if cotizar:

    if not coberturas_config:
        st.warning("⚠️ Activa al menos una cobertura con Suma Asegurada > 0.")
        st.stop()

    pagos_map = {"Anual": 1, "Semestral": 2, "Trimestral": 4, "Mensual": 12}
    m  = pagos_map[forma_pago]
    ff = factor_financiero(TASA_ANUAL, m) if m > 1 else 1.0

    filas = []; sin_datos = []
    for cfg in coberturas_config:
        cob = cfg["cobertura"]; sa = cfg["sa"]
        cr_g = get_cr_global(cr_global_df, cob, tipo_bien)
        if cr_g is None: sin_datos.append(cob); continue
        f_geo  = get_f_geo(factor_geo_df, cob, tipo_bien, estado)
        f_ded  = get_f_ded(factor_ded_df, deducible)
        cr_fin = cr_g * f_geo * f_ded
        pr     = sa * cr_fin / 1000
        pt     = pr / (1 - CARGOS)
        prima  = pt * ff if m > 1 else pt
        filas.append({
            "cve": TIPOS_ORDEN.get(tipo_bien, ""), "tipo_bien": tipo_bien,
            "cobertura": cob, "sa": sa, "cr_global": cr_g, "f_geo": f_geo,
            "cr_final": cr_fin, "pr": pr, "deducible": deducible,
            "f_ded": f_ded, "prima": prima, "recibo": prima / m, "pt": pt,
        })

    if sin_datos:
        st.warning(f"⚠️ Sin datos para: {', '.join(sin_datos)}. Se excluyen.")
    if not filas:
        st.error("Sin datos para ninguna cobertura seleccionada."); st.stop()

    total_sa    = sum(f["sa"]    for f in filas)
    total_pr    = sum(f["pr"]    for f in filas)
    total_pt    = sum(f["pt"]    for f in filas)
    total_prima = sum(f["prima"] for f in filas)
    total_recibo= sum(f["recibo"] for f in filas)
    total_cr    = total_pr / total_sa * 1000 if total_sa > 0 else 0
    recargo_pct = (ff - 1) * 100 if m > 1 else 0.0
    label_recibo = {1:"Pago Único Anual", 2:"Pago Semestral",
                    4:"Pago Trimestral", 12:"Pago Mensual"}[m]

    # ── Datos de póliza ──
    st.markdown('<div class="seccion-titulo">Datos de la Póliza</div>', unsafe_allow_html=True)
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("No. Póliza",      no_poliza)
    p2.metric("Fecha Emisión",   fecha_emision.strftime("%d/%m/%Y"))
    p3.metric("Inicio Vigencia", fecha_inicio.strftime("%d/%m/%Y"))
    p4.metric("Fin Vigencia",    fecha_fin.strftime("%d/%m/%Y"))

    # ── Datos del riesgo ──
    st.markdown('<div class="seccion-titulo">Datos del Asegurado</div>', unsafe_allow_html=True)
    nombre       = st.text_input("Nombre(s)")
    ap_paterno   = st.text_input("Apellido Paterno")
    ap_materno   = st.text_input("Apellido Materno")
    correo       = st.text_input("Correo electrónico")
    telefono     = st.text_input("Teléfono celular")

    st.markdown('<div class="seccion-titulo">Datos del Riesgo</div>', unsafe_allow_html=True)
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Estado",      estado)
    r2.metric("Tipo de Bien",tipo_bien)
    r3.metric("Deducible",   f"${deducible:,.0f}")
    r4.metric("Coberturas",  str(len(filas)))

    # ── Tabla de cotización ──
    st.markdown('<div class="seccion-titulo">Detalle de Coberturas</div>', unsafe_allow_html=True)

    filas_html = ""
    for f in filas:
        filas_html += (
            "<tr>"
            f"<td>{f['tipo_bien']}</td>"
            f"<td>{f['cobertura']}</td>"
            f"<td><span class='badge-si'>SI</span></td>"
            f"<td>{fmt_m(f['sa'])}</td>"
            f"<td>{f['cr_final']:.5f}</td>"
            f"<td>{fmt_m(f['pr'])}</td>"
            f"<td><span class='badge-si'>SI</span></td>"
            f"<td>{fmt_m(f['deducible'])}</td>"
            f"<td>{f['f_ded']:.4f}</td>"
            f"<td style='background:#fdf0ee;color:#7B1C35;font-weight:700;text-align:right;padding:0.55rem 0.8rem'>{fmt_m(f['prima'])}</td>"
            "</tr>"
        )

    tabla_html = (
        "<div class='card' style='overflow-x:auto;padding:0.5rem 0'>"
        "<table class='tbl'>"
        "<thead><tr>"
        "<th>Tipo de Bien</th><th>Cobertura</th><th>SI/NO</th>"
        "<th>Suma Asegurada</th><th>Cuota Riesgo ‰</th><th>Prima de Riesgo</th>"
        "<th>Deducible</th><th>Valor Deducible</th><th>f<sub>DED</sub></th><th style='background:#5a1326;color:#f5c8b8;font-weight:700;text-align:right;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;padding:0.65rem 0.8rem'>Prima</th>"
        "</tr></thead><tbody>"
        + filas_html +
        "<tr class='total-row'>"
        "<td colspan='3'>TOTAL</td>"
        f"<td>{fmt_m(total_sa)}</td>"
        f"<td>{total_cr:.5f}</td>"
        f"<td>{fmt_m(total_pr)}</td>"
        "<td colspan='2'></td>"
        "<td>Total Prima</td>"
        f"<td style='background:#7B1C35;color:#f5ede8;font-weight:700;text-align:right;font-size:0.95rem;padding:0.55rem 0.8rem'>{fmt_m(total_prima)}</td>"
        "</tr></tbody></table></div>"
    )
    st.markdown(tabla_html, unsafe_allow_html=True)

    # ── Tarjetas resumen ──
    st.markdown('<div class="seccion-titulo">Resumen de Prima y Pago</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"""
        <div class="card">
            <div class="card-label">Prima Total (con recargo financiero)</div>
            <div class="card-value">{fmt_m(total_prima)}</div>
            <div class="card-sub">Adm {GASTO_ADMIN*100:.0f}% · Adq {GASTO_ADQ*100:.0f}% · Utilidad {MARGEN_UT*100:.0f}%</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="card">
            <div class="card-label">{label_recibo}</div>
            <div class="card-value">{fmt_m(total_recibo)}</div>
            <div class="card-sub">{m} pago{'s' if m>1 else ''} · Recargo {recargo_pct:.2f}% · Tasa {TASA_ANUAL*100:.0f}% anual</div>
        </div>
        """, unsafe_allow_html=True)

    # Desglose numérico
    st.markdown(f"""
    <div class="card">
        <div class="desglose-row">
            <span class="desglose-label">Prima de Riesgo</span>
            <span class="desglose-valor">{fmt_m(total_pr)}</span>
        </div>
        <div class="desglose-row">
            <span class="desglose-label">Prima de Tarifa (÷ {1-CARGOS:.2f})</span>
            <span class="desglose-valor">{fmt_m(total_pt)}</span>
        </div>
        <div class="desglose-row">
            <span class="desglose-label">Recargo financiero · f = {ff:.6f}</span>
            <span class="desglose-valor">{fmt_m(total_prima - total_pt)}</span>
        </div>
        <div class="desglose-total">
            <span>Prima Total con Recargo</span>
            <span>{fmt_m(total_prima)}</span>
        </div>
        <div class="desglose-total">
            <span>{label_recibo} (÷{m})</span>
            <span>{fmt_m(total_recibo)}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Datos del asegurado en resultados ──
    nombre_completo = f"{nombre} {ap_paterno} {ap_materno}".strip() or "—"
    st.markdown('<div class="seccion-titulo">Datos del Asegurado</div>', unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)
    a1.metric("Nombre", nombre_completo)
    a2.metric("Correo", correo or "-")
    a3.metric("Teléfono", telefono or "-")

    st.markdown(f"""
    <div class="alerta">
        <strong>Póliza {no_poliza}</strong> · Vigencia {fecha_inicio.strftime("%d/%m/%Y")} – {fecha_fin.strftime("%d/%m/%Y")} ·
        Estado: {estado} · Deducible: ${deducible:,.0f}<br>
        <em>Cotización informativa basada en estadística SESA 2020–2024. No constituye oferta de contrato.</em>
    </div>
    """, unsafe_allow_html=True)

    # ── Contacto Bouclier ──
    st.markdown("""
    <div class="card" style="margin-top:1.5rem;border-left:4px solid #7B1C35;padding:1rem 1.5rem">
        <div style="font-family:'Cormorant Garamond',serif;font-size:1.1rem;font-weight:700;color:#7B1C35;margin-bottom:0.6rem">
            Contacto Bouclier Seguros de Daños
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:0.4rem 2rem;font-size:0.88rem;color:#4a4a4a">
            <div>📞 <strong>Atención a Clientes:</strong> 800 268 2543</div>
            <div>📧 <strong>Correo:</strong> contacto@bouclier.com.mx</div>
            <div>📱 <strong>WhatsApp:</strong> +52 55 4768 2300</div>
            <div>🌐 <strong>Web:</strong> www.bouclier.com.mx</div>
        </div>
        <div style="font-size:0.78rem;color:#888;margin-top:0.6rem">
            Lunes a Viernes 9:00 – 18:00 h · Ciudad de México
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Descarga PDF ──
    if FPDF_OK:
        def generar_pdf():
            pdf = FPDF()
            pdf.add_page()
            pdf.set_margins(18, 18, 18)

            # Encabezado
            pdf.set_fill_color(123, 28, 53)
            pdf.rect(0, 0, 210, 28, 'F')
            pdf.set_text_color(245, 237, 232)
            pdf.set_font("Helvetica", "B", 16)
            pdf.set_xy(18, 8)
            pdf.cell(0, 10, "BOUCLIER SEGUROS DE DAÑOS", ln=False)
            pdf.set_font("Helvetica", "", 9)
            pdf.set_xy(18, 18)
            pdf.cell(0, 6, "Cotizacion - Casa Habitacion - Ramo Incendio", ln=True)

            pdf.set_text_color(30, 30, 30)
            pdf.set_xy(18, 34)

            def titulo(txt):
                pdf.set_font("Helvetica", "B", 10)
                pdf.set_text_color(123, 28, 53)
                pdf.cell(0, 7, txt, ln=True)
                pdf.set_draw_color(196, 181, 165)
                pdf.line(18, pdf.get_y(), 192, pdf.get_y())
                pdf.ln(2)
                pdf.set_text_color(30, 30, 30)

            def fila(lbl, val):
                pdf.set_font("Helvetica", "", 9)
                pdf.set_text_color(100, 100, 100)
                pdf.cell(55, 6, lbl, ln=False)
                pdf.set_text_color(30, 30, 30)
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(0, 6, str(val), ln=True)

            # Datos póliza
            titulo("DATOS DE LA PÓLIZA")
            fila("No. Póliza:", no_poliza)
            fila("Fecha Emisión:", fecha_emision.strftime("%d/%m/%Y"))
            fila("Inicio Vigencia:", fecha_inicio.strftime("%d/%m/%Y"))
            fila("Fin Vigencia:", fecha_fin.strftime("%d/%m/%Y"))
            pdf.ln(3)

            # Datos asegurado
            titulo("DATOS DEL ASEGURADO")
            fila("Nombre:", f"{nombre} {ap_paterno} {ap_materno}".strip().replace("-","-") or "-")
            fila("Correo:", correo or "-")
            fila("Teléfono:", telefono or "-")
            pdf.ln(3)

            # Datos riesgo
            titulo("DATOS DEL RIESGO")
            fila("Entidad Federativa:", estado)
            fila("Tipo de Bien:", tipo_bien)
            fila("Deducible:", f"${deducible:,.0f}")
            pdf.ln(3)

            # Tabla coberturas
            titulo("DETALLE DE COBERTURAS")
            pdf.set_fill_color(123, 28, 53)
            pdf.set_text_color(245, 237, 232)
            pdf.set_font("Helvetica", "B", 8)
            cols = [("Cobertura",52),("Suma Asegurada",38),("Cuota o/oo",24),("Prima Riesgo",30),("Prima Total",30)]
            for h,w in cols:
                pdf.cell(w, 7, h, border=0, align="C", fill=True)
            pdf.ln()
            pdf.set_text_color(30, 30, 30)
            for i, f_row in enumerate(filas):
                fill = i % 2 == 0
                pdf.set_fill_color(253, 240, 238) if fill else pdf.set_fill_color(255, 255, 255)
                pdf.set_font("Helvetica", "", 8)
                pdf.cell(52, 6, f_row["cobertura"][:28], border=0, fill=fill)
                pdf.cell(38, 6, f"${f_row['sa']:,.0f}", border=0, align="R", fill=fill)
                pdf.cell(24, 6, f"{f_row['cr_final']:.5f}", border=0, align="R", fill=fill)
                pdf.cell(30, 6, f"${f_row['pr']:,.2f}", border=0, align="R", fill=fill)
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(123, 28, 53)
                pdf.cell(30, 6, f"${f_row['prima']:,.2f}", border=0, align="R", fill=fill, ln=True)
                pdf.set_text_color(30, 30, 30)

            # Total
            pdf.set_fill_color(123, 28, 53)
            pdf.set_text_color(245, 237, 232)
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(52+38+24+30, 7, "TOTAL", border=0, fill=True)
            pdf.cell(30, 7, f"${total_prima:,.2f}", border=0, align="R", fill=True, ln=True)
            pdf.ln(4)

            # Resumen
            titulo("RESUMEN DE PRIMA")
            fila("Prima de Riesgo:", f"${total_pr:,.2f}")
            fila("Prima de Tarifa:", f"${total_pt:,.2f}")
            fila("Forma de Pago:", forma_pago)
            fila(f"{label_recibo}:", f"${total_recibo:,.2f}")
            fila("Prima Total:", f"${total_prima:,.2f}")
            pdf.ln(4)

            # Pie con contacto
            pdf.set_fill_color(245, 237, 232)
            pdf.set_x(18)
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_text_color(123, 28, 53)
            pdf.cell(0, 6, "Contacto Bouclier Seguros de Daños", ln=True)
            pdf.set_font("Helvetica", "", 8)
            pdf.set_text_color(80, 80, 80)
            pdf.cell(0, 5, "800 268 2543  |  contacto@bouclier.com.mx  |  +52 55 4768 2300  |  www.bouclier.com.mx", ln=True)
            pdf.cell(0, 5, "Cotización informativa. No constituye oferta de contrato.", ln=True)

            return bytes(pdf.output())

        pdf_bytes = generar_pdf()
        st.download_button(
            label="⬇️  Descargar Cotización PDF",
            data=pdf_bytes,
            file_name=f"Cotizacion_{no_poliza}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    else:
        st.info("Instala fpdf2 para habilitar la descarga: `pip install fpdf2`")

else:
    # Pantalla vacía
    st.markdown("""
    <div style="text-align:center;padding:5rem 2rem">
        <div style="margin-bottom:1rem"><svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" style="width:80px;height:80px;display:inline-block">
  <path d="M100 185 C60 170 28 140 20 100 L20 45 C20 38 26 33 33 35 L100 18 L167 35 C174 33 180 38 180 45 L180 100 C172 140 140 170 100 185Z" fill="#4a0d1a" opacity="0.3" transform="translate(3,4)"/>
  <path d="M100 180 C62 165 30 136 22 98 L22 46 C22 40 27 36 33 37 L100 21 L167 37 C173 36 178 40 178 46 L178 98 C170 136 138 165 100 180Z" fill="#7B1C35"/>
  <path d="M100 170 C68 156 40 130 33 96 L33 52 C33 48 37 45 41 46 L100 32 L159 46 C163 45 167 48 167 52 L167 96 C160 130 132 156 100 170Z" fill="none" stroke="#C4B5A5" stroke-width="1.5"/>
  <line x1="35" y1="100" x2="165" y2="100" stroke="#C4B5A5" stroke-width="1" opacity="0.6"/>
  <line x1="100" y1="34" x2="100" y2="168" stroke="#C4B5A5" stroke-width="1" opacity="0.6"/>
  <circle cx="67" cy="66" r="14" fill="#C4B5A5" opacity="0.25"/>
  <circle cx="67" cy="66" r="8" fill="#C4B5A5" opacity="0.4"/>
  <circle cx="67" cy="66" r="3" fill="#f5ede8"/>
  <path d="M122 52 L133 66 L122 80 L111 66 Z" fill="#C4B5A5" opacity="0.35"/>
  <path d="M122 57 L129 66 L122 75 L115 66 Z" fill="#f5ede8" opacity="0.5"/>
  <path d="M55 108 L79 108 L79 132 L55 132 Z" fill="none" stroke="#C4B5A5" stroke-width="1.2" opacity="0.5"/>
  <path d="M60 113 L74 113 L74 127 L60 127 Z" fill="#C4B5A5" opacity="0.3"/>
  <path d="M109 115 L121 109 L133 115 L133 129 L121 135 L109 129 Z" fill="none" stroke="#C4B5A5" stroke-width="1.2" opacity="0.5"/>
  <circle cx="121" cy="122" r="4" fill="#C4B5A5" opacity="0.5"/>
  <path d="M94 165 L100 178 L106 165 Z" fill="#C4B5A5" opacity="0.6"/>
  <circle cx="45" cy="47" r="3" fill="#C4B5A5" opacity="0.7"/>
  <circle cx="155" cy="47" r="3" fill="#C4B5A5" opacity="0.7"/>
</svg></div>
        <div style="font-family:'Cormorant Garamond',serif;font-size:1.8rem;color:#7B1C35;margin-bottom:0.5rem;font-weight:700">
            Configura los parámetros en el panel izquierdo
        </div>
        <div style="font-size:0.88rem;color:#9a8a84">
            Póliza · Estado · Tipo de bien · Coberturas con su Suma Asegurada · Deducible · Forma de pago<br>
            y presiona <strong style="color:#7B1C35">Calcular Cotización</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)


    