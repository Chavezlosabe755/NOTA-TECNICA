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

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #0d0d0d !important;
    border-right: none;
}
[data-testid="stSidebar"] * { color: #f5ede8 !important; }
[data-testid="stSidebar"] .seccion-titulo {
    color: #C4B5A5 !important;
    border-bottom-color: rgba(196,181,165,0.3) !important;
}
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] select,
[data-testid="stSidebar"] .stSelectbox > div,
[data-testid="stSidebar"] .stNumberInput > div > div {
    background-color: rgba(255,255,255,0.08) !important;
    border-color: rgba(196,181,165,0.4) !important;
    color: #f5ede8 !important;
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
[data-testid="stSidebar"] label { color: #f0e6df !important; font-size: 0.82rem !important; }
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
    background: #f5ede8;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    flex-shrink: 0;
    overflow: hidden;
    border: 3px solid #C4B5A5;
}
.logo-wrap svg { width: 70px; height: 70px; }

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
.tbl td {
    padding: 0.55rem 0.8rem;
    border-bottom: 1px solid #f0e8e3;
    color: #2a1a1e;
    text-align: right;
    vertical-align: middle;
}
.tbl td:first-child, .tbl td:nth-child(2), .tbl td:nth-child(3) { text-align: left; }
.tbl tr:hover td { background: #fdf5f2; }
.tbl tr.total-row td {
    background: #fdf5f2;
    color: #7B1C35;
    font-weight: 700;
    border-top: 2px solid #7B1C35;
    border-bottom: none;
    font-size: 0.88rem;
}
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
    background: rgba(255,255,255,0.15);
    display: flex; align-items: center; justify-content: center;
    border: 2px solid rgba(196,181,165,0.5); flex-shrink: 0;
}
.sidebar-logo-wrap .sb-logo svg { width: 36px; height: 36px; }
.sidebar-brand { font-family: 'Cormorant Garamond', serif; font-size: 1.1rem; font-weight: 700; color: #C4B5A5 !important; }
.sidebar-sub { font-size: 0.68rem; color: #C4B5A5 !important; letter-spacing: 0.1em; text-transform: uppercase; }


/* Sidebar negro — mejor visibilidad */
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stDateInput input {
    background-color: #1a1a1a !important;
    border-color: #333 !important;
    color: #e0d0c8 !important;
}
[data-testid="stSidebar"] .stSelectbox > div > div,
[data-testid="stSidebar"] .stNumberInput input {
    background-color: #1a1a1a !important;
    border-color: #333 !important;
    color: #e0d0c8 !important;
}
[data-testid="stSidebar"] .stRadio > div { gap: 0.3rem; }

footer { visibility: hidden; }
#MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# SVG del logo Bouclier (águila/fénix estilizado con círculo y texto)
LOGO_SVG = """
<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
  <circle cx="50" cy="50" r="48" fill="#7B1C35" stroke="#C4B5A5" stroke-width="2"/>
  <circle cx="50" cy="50" r="42" fill="none" stroke="#C4B5A5" stroke-width="0.8"/>
  <!-- Cuerpo del ave -->
  <ellipse cx="50" cy="54" rx="10" ry="14" fill="#f5ede8"/>
  <!-- Alas -->
  <path d="M40 50 Q25 38 22 48 Q30 52 40 55 Z" fill="#f5ede8"/>
  <path d="M60 50 Q75 38 78 48 Q70 52 60 55 Z" fill="#f5ede8"/>
  <!-- Alas secundarias -->
  <path d="M40 46 Q28 32 26 42 Q33 46 40 50 Z" fill="#C4B5A5"/>
  <path d="M60 46 Q72 32 74 42 Q67 46 60 50 Z" fill="#C4B5A5"/>
  <!-- Cabeza -->
  <circle cx="50" cy="40" r="8" fill="#f5ede8"/>
  <!-- Pico -->
  <path d="M50 44 L53 47 L50 46 Z" fill="#C4B5A5"/>
  <!-- Ojo -->
  <circle cx="52" cy="39" r="1.5" fill="#7B1C35"/>
  <!-- Cola -->
  <path d="M44 66 Q50 72 56 66 Q53 70 50 74 Q47 70 44 66 Z" fill="#f5ede8"/>
  <!-- Plumas corona -->
  <path d="M46 33 Q47 26 50 28 Q53 26 54 33" fill="none" stroke="#C4B5A5" stroke-width="1.2"/>
  <circle cx="50" cy="26" r="1.5" fill="#C4B5A5"/>
  <!-- Texto BOUCLIER -->
  <path id="curve" d="M 15 50 A 35 35 0 0 1 85 50" fill="none"/>
  <text font-family="serif" font-size="7" fill="#C4B5A5" letter-spacing="2">
    <textPath href="#curve" startOffset="15%">BOUCLIER</textPath>
  </text>
</svg>
"""

LOGO_SVG_SMALL = """
<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
  <circle cx="50" cy="50" r="48" fill="none"/>
  <ellipse cx="50" cy="54" rx="10" ry="14" fill="#f5ede8"/>
  <path d="M40 50 Q25 38 22 48 Q30 52 40 55 Z" fill="#f5ede8"/>
  <path d="M60 50 Q75 38 78 48 Q70 52 60 55 Z" fill="#f5ede8"/>
  <path d="M40 46 Q28 32 26 42 Q33 46 40 50 Z" fill="#C4B5A5"/>
  <path d="M60 46 Q72 32 74 42 Q67 46 60 50 Z" fill="#C4B5A5"/>
  <circle cx="50" cy="40" r="8" fill="#f5ede8"/>
  <path d="M50 44 L53 47 L50 46 Z" fill="#C4B5A5"/>
  <circle cx="52" cy="39" r="1.5" fill="#7B1C35"/>
  <path d="M44 66 Q50 72 56 66 Q53 70 50 74 Q47 70 44 66 Z" fill="#f5ede8"/>
  <path d="M46 33 Q47 26 50 28 Q53 26 54 33" fill="none" stroke="#C4B5A5" stroke-width="1.2"/>
  <circle cx="50" cy="26" r="1.5" fill="#C4B5A5"/>
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
        <div class="sb-logo">{LOGO_SVG_SMALL}</div>
        <div>
            <div class="sidebar-brand">Bouclier</div>
            <div class="sidebar-sub">Seguros de Daños</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="seccion-titulo">Datos de la Póliza</div>', unsafe_allow_html=True)
    no_poliza     = st.text_input("No. Póliza", value="INC0001")
    fecha_emision = st.date_input("Fecha de Emisión", value=date.today(), format="DD/MM/YYYY")
    fecha_inicio  = st.date_input("Fecha Inicio Vigencia", value=date.today(), format="DD/MM/YYYY")
    fecha_fin     = st.date_input("Fecha Fin Vigencia", format="DD/MM/YYYY",
                                   value=date.today().replace(year=date.today().year + 1))

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
                step=10_000.0, format="%.2f",
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
        <div class="logo-wrap">{LOGO_SVG}</div>
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
            f"<td>{f['cve']}</td>"
            f"<td>{f['tipo_bien']}</td>"
            f"<td>{f['cobertura']}</td>"
            f"<td><span class='badge-si'>SI</span></td>"
            f"<td>{fmt_m(f['sa'])}</td>"
            f"<td>{f['cr_final']:.5f}</td>"
            f"<td>{fmt_m(f['pr'])}</td>"
            f"<td><span class='badge-si'>SI</span></td>"
            f"<td>{fmt_m(f['deducible'])}</td>"
            f"<td>{f['f_ded']:.4f}</td>"
            f"<td><strong>{fmt_m(f['prima'])}</strong></td>"
            "</tr>"
        )

    tabla_html = (
        "<div class='card' style='overflow-x:auto;padding:0.5rem 0'>"
        "<table class='tbl'>"
        "<thead><tr>"
        "<th>Cve</th><th>Tipo de Bien</th><th>Cobertura</th><th>SI/NO</th>"
        "<th>Suma Asegurada</th><th>Cuota Riesgo ‰</th><th>Prima de Riesgo</th>"
        "<th>Deducible</th><th>Valor Deducible</th><th>f<sub>DED</sub></th><th>Prima</th>"
        "</tr></thead><tbody>"
        + filas_html +
        "<tr class='total-row'>"
        "<td colspan='4'>TOTAL</td>"
        f"<td>{fmt_m(total_sa)}</td>"
        f"<td>{total_cr:.5f}</td>"
        f"<td>{fmt_m(total_pr)}</td>"
        "<td colspan='2'></td>"
        "<td>Total Prima</td>"
        f"<td>{fmt_m(total_prima)}</td>"
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

    st.markdown(f"""
    <div class="alerta">
        <strong>Póliza {no_poliza}</strong> · Vigencia {fecha_inicio.strftime("%d/%m/%Y")} – {fecha_fin.strftime("%d/%m/%Y")} ·
        Estado: {estado} · Deducible: ${deducible:,.0f}<br>
        <em>Cotización informativa basada en estadística SESA 2020–2024. No constituye oferta de contrato.</em>
    </div>
    """, unsafe_allow_html=True)

else:
    # Pantalla vacía
    st.markdown("""
    <div style="text-align:center;padding:5rem 2rem">
        <div style="font-size:3.5rem;margin-bottom:1rem">🛡️</div>
        <div style="font-family:'Cormorant Garamond',serif;font-size:1.8rem;color:#7B1C35;margin-bottom:0.5rem;font-weight:700">
            Configura los parámetros en el panel izquierdo
        </div>
        <div style="font-size:0.88rem;color:#9a8a84">
            Póliza · Estado · Tipo de bien · Coberturas con su Suma Asegurada · Deducible · Forma de pago<br>
            y presiona <strong style="color:#7B1C35">Calcular Cotización</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)