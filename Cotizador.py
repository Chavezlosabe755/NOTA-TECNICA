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

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Cotizador | Casa Habitación",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Fondo general */
.stApp {
    background-color: #0f1117;
    color: #e8e8e8;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #161b27;
    border-right: 1px solid #2a3142;
}

/* Título principal */
.titulo-principal {
    font-family: 'DM Serif Display', serif;
    font-size: 2.4rem;
    color: #f5f0e8;
    line-height: 1.15;
    margin-bottom: 0.2rem;
}
.subtitulo {
    font-size: 0.95rem;
    color: #7a8599;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 2rem;
}

/* Tarjetas de resultado */
.card {
    background: #1a2035;
    border: 1px solid #2a3142;
    border-radius: 12px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1rem;
}
.card-label {
    font-size: 0.78rem;
    color: #7a8599;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom: 0.3rem;
}
.card-value {
    font-family: 'DM Serif Display', serif;
    font-size: 2rem;
    color: #d4b896;
    line-height: 1.1;
}
.card-sub {
    font-size: 0.82rem;
    color: #7a8599;
    margin-top: 0.2rem;
}

/* Tarjeta desglose */
.desglose-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.5rem 0;
    border-bottom: 1px solid #2a3142;
    font-size: 0.9rem;
}
.desglose-row:last-child { border-bottom: none; }
.desglose-label { color: #a0aabb; }
.desglose-valor { color: #e8e8e8; font-weight: 500; }

/* Alerta */
.alerta {
    background: #1e1a2e;
    border-left: 3px solid #d4b896;
    border-radius: 6px;
    padding: 0.8rem 1rem;
    font-size: 0.85rem;
    color: #a0aabb;
    margin-top: 1rem;
}

/* Sección */
.seccion-titulo {
    font-size: 0.72rem;
    color: #7a8599;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin: 1.4rem 0 0.6rem 0;
    border-bottom: 1px solid #2a3142;
    padding-bottom: 0.4rem;
}

/* Ocultar footer de Streamlit */
footer { visibility: hidden; }
#MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
EXCEL_PATH = "RESULTADOS_NOTA_TECNICA.xlsx"

@st.cache_data(show_spinner="Cargando tablas actuariales…")
def cargar_datos(path):
    xls = pd.ExcelFile(path)
    cr_global      = pd.read_excel(xls, "CR_Global")
    factor_geo     = pd.read_excel(xls, "Factor_Geografico")
    factor_ded     = pd.read_excel(xls, "Factor_Deducible")
    prima_tarifa   = pd.read_excel(xls, "Prima_Tarifa")
    return cr_global, factor_geo, factor_ded, prima_tarifa


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt_moneda(valor):
    if valor is None or np.isnan(valor):
        return "—"
    return f"${valor:,.2f}"

def fmt_millar(valor):
    if valor is None or np.isnan(valor):
        return "—"
    return f"{valor:.6f} ‰"

def factor_financiero(j, m):
    """Factor de recargo por pago fraccionado."""
    denominador = sum((1 + j / m) ** (-k) for k in range(m))
    return m / denominador

def calcular_cotizacion(
    cr_global_df,
    factor_geo_df,
    factor_ded_df,
    prima_tarifa_df,
    cobertura,
    tipo_bien,
    estado,
    suma_asegurada,
    deducible,
    gasto_admin,
    gasto_adq,
    margen_ut,
    tasa_anual,
    forma_pago,   # "Anual", "Semestral", "Trimestral", "Mensual"
):
    pagos_map = {"Anual": 1, "Semestral": 2, "Trimestral": 4, "Mensual": 12}
    m = pagos_map[forma_pago]

    # 1. CR Global
    mask_g = (
        (cr_global_df["COBERTURA"]    == cobertura) &
        (cr_global_df["TIPO DE BIEN"] == tipo_bien)
    )
    fila_g = cr_global_df[mask_g]
    if fila_g.empty:
        return None, "No hay cuota global para esta combinación de cobertura y tipo de bien."
    cr_global = fila_g["CR_GLOBAL"].values[0]

    if pd.isna(cr_global) or cr_global == 0:
        return None, "La cuota global para esta combinación es cero o nula."

    # 2. Factor geográfico
    mask_geo = (
        (factor_geo_df["COBERTURA"]    == cobertura) &
        (factor_geo_df["TIPO DE BIEN"] == tipo_bien) &
        (factor_geo_df["ENTIDAD"]      == estado)
    )
    fila_geo = factor_geo_df[mask_geo]
    f_geo = fila_geo["FACTOR_GEO"].values[0] if not fila_geo.empty else 1.0
    if pd.isna(f_geo) or f_geo <= 0:
        f_geo = 1.0

    # 3. Factor por deducible
    mask_ded = factor_ded_df["DEDUCIBLE"] == deducible
    fila_ded = factor_ded_df[mask_ded]
    f_ded = fila_ded["f"].values[0] if not fila_ded.empty else 1.0
    if pd.isna(f_ded):
        f_ded = 1.0

    # 4. Cuota de riesgo final (al millar)
    cr_final = cr_global * f_geo * f_ded

    # 5. Prima de riesgo
    prima_riesgo = suma_asegurada * cr_final / 1000

    # 6. Prima de tarifa
    cargos = gasto_admin + gasto_adq + margen_ut
    if cargos >= 1:
        return None, "Los gastos totales no pueden ser ≥ 100%."
    prima_tarifa = prima_riesgo / (1 - cargos)

    # 7. Recargo financiero
    ff = factor_financiero(tasa_anual, m) if m > 1 else 1.0
    prima_con_recargo = prima_tarifa * ff if m > 1 else prima_tarifa
    recibo = prima_con_recargo / m

    # 8. Desglose de gastos
    g_admin  = prima_tarifa * gasto_admin
    g_adq    = prima_tarifa * gasto_adq
    g_ut     = prima_tarifa * margen_ut
    recargo_pct = (ff - 1) * 100 if m > 1 else 0.0

    return {
        "cr_global"        : cr_global,
        "f_geo"            : f_geo,
        "f_ded"            : f_ded,
        "cr_final"         : cr_final,
        "prima_riesgo"     : prima_riesgo,
        "prima_tarifa"     : prima_tarifa,
        "prima_con_recargo": prima_con_recargo,
        "recibo"           : recibo,
        "g_admin"          : g_admin,
        "g_adq"            : g_adq,
        "g_ut"             : g_ut,
        "recargo_pct"      : recargo_pct,
        "m"                : m,
        "ff"               : ff,
    }, None


# ─────────────────────────────────────────────
# ENCABEZADO
# ─────────────────────────────────────────────
col_logo, col_titulo = st.columns([1, 5])
with col_titulo:
    st.markdown('<div class="titulo-principal">Cotizador<br>Casa Habitación</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitulo">Bouclier Seguros de Daños · Ramo Incendio</div>', unsafe_allow_html=True)

st.divider()

# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────
if not os.path.exists(EXCEL_PATH):
    st.error(
        f"⚠️ No se encontró el archivo **{EXCEL_PATH}** en el directorio actual.\n\n"
        "Asegúrate de correr primero el notebook y de colocar el Excel en la misma carpeta que este script."
    )
    st.stop()

try:
    cr_global_df, factor_geo_df, factor_ded_df, prima_tarifa_df = cargar_datos(EXCEL_PATH)
except Exception as e:
    st.error(f"Error al leer el Excel: {e}")
    st.stop()

# Catálogos disponibles en los datos
coberturas_disponibles = sorted(cr_global_df["COBERTURA"].dropna().unique().tolist())
tipos_bien_disponibles = sorted(cr_global_df["TIPO DE BIEN"].dropna().unique().tolist())
estados_disponibles    = sorted(factor_geo_df["ENTIDAD"].dropna().unique().tolist())
deducibles_disponibles = sorted(factor_ded_df["DEDUCIBLE"].dropna().unique().tolist())


# ─────────────────────────────────────────────
# SIDEBAR — PARÁMETROS DE COTIZACIÓN
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="seccion-titulo">Datos del Riesgo</div>', unsafe_allow_html=True)

    estado = st.selectbox("Entidad Federativa", estados_disponibles)

    tipo_bien = st.selectbox(
        "Tipo de Bien",
        tipos_bien_disponibles,
        help="Contenidos, Edificio o Contenidos y Edificio"
    )

    cobertura = st.selectbox(
        "Cobertura",
        coberturas_disponibles,
        help="Selecciona la cobertura a contratar"
    )

    suma_asegurada = st.number_input(
        "Suma Asegurada ($)",
        min_value=10_000.0,
        max_value=500_000_000.0,
        value=1_000_000.0,
        step=10_000.0,
        format="%.2f",
        help="Valor asegurado en pesos mexicanos"
    )

    deducible = st.selectbox(
        "Deducible ($)",
        deducibles_disponibles,
        index=0,
        format_func=lambda x: f"${x:,.0f}",
        help="Monto que absorbe el asegurado por siniestro"
    )

    st.markdown('<div class="seccion-titulo">Forma de Pago</div>', unsafe_allow_html=True)

    forma_pago = st.radio(
        "Periodicidad",
        ["Anual", "Semestral", "Trimestral", "Mensual"],
        horizontal=False,
    )

    st.markdown('<div class="seccion-titulo">Supuestos Técnicos</div>', unsafe_allow_html=True)

    with st.expander("Gastos y tasa financiera", expanded=False):
        gasto_admin = st.slider(
            "Gastos de Administración",
            0.0, 0.50, 0.20, 0.01,
            format="%.0f%%",
            help="% sobre prima de tarifa"
        )
        gasto_adq = st.slider(
            "Gastos de Adquisición",
            0.0, 0.30, 0.05, 0.01,
            format="%.0f%%",
        )
        margen_ut = st.slider(
            "Margen de Utilidad",
            0.0, 0.30, 0.06, 0.01,
            format="%.0f%%",
        )
        tasa_anual = st.number_input(
            "Tasa Anual Financiera (%)",
            min_value=0.0, max_value=50.0,
            value=8.0, step=0.5,
            format="%.1f"
        ) / 100

    cotizar = st.button("Calcular Cotización", use_container_width=True, type="primary")


# ─────────────────────────────────────────────
# ÁREA PRINCIPAL
# ─────────────────────────────────────────────
if cotizar:
    resultado, error = calcular_cotizacion(
        cr_global_df, factor_geo_df, factor_ded_df, prima_tarifa_df,
        cobertura, tipo_bien, estado,
        suma_asegurada, deducible,
        gasto_admin, gasto_adq, margen_ut,
        tasa_anual, forma_pago,
    )

    if error:
        st.warning(f"⚠️ {error}")
    else:
        r = resultado

        # ── Resumen de inputs ──
        st.markdown('<div class="seccion-titulo">Resumen del Riesgo</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Estado", estado)
        c2.metric("Tipo de Bien", tipo_bien)
        c3.metric("Cobertura", cobertura)
        c4.metric("Suma Asegurada", fmt_moneda(suma_asegurada))

        st.markdown('<div class="seccion-titulo">Resultado de la Cotización</div>', unsafe_allow_html=True)

        # ── Tarjetas principales ──
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"""
            <div class="card">
                <div class="card-label">Prima de Tarifa Anual</div>
                <div class="card-value">{fmt_moneda(r['prima_con_recargo'])}</div>
                <div class="card-sub">Incluye recargo financiero</div>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            label_pago = {1:"Pago Único", 2:"Recibo Semestral", 4:"Recibo Trimestral", 12:"Recibo Mensual"}[r['m']]
            st.markdown(f"""
            <div class="card">
                <div class="card-label">{label_pago}</div>
                <div class="card-value">{fmt_moneda(r['recibo'])}</div>
                <div class="card-sub">{r['m']} pago{'s' if r['m']>1 else ''} de igual valor</div>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div class="card">
                <div class="card-label">Cuota de Riesgo Final</div>
                <div class="card-value">{r['cr_final']:.6f} ‰</div>
                <div class="card-sub">Por cada $1,000 de suma asegurada</div>
            </div>
            """, unsafe_allow_html=True)

        # ── Desglose técnico ──
        st.markdown('<div class="seccion-titulo">Desglose Técnico</div>', unsafe_allow_html=True)

        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("**Construcción de la Cuota de Riesgo**")
            st.markdown(f"""
            <div class="card">
                <div class="desglose-row">
                    <span class="desglose-label">CR Global (promedio 2020-2024)</span>
                    <span class="desglose-valor">{r['cr_global']:.6f} ‰</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">× Factor Geográfico ({estado})</span>
                    <span class="desglose-valor">{r['f_geo']:.6f}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">× Factor Deducible (${deducible:,.0f})</span>
                    <span class="desglose-valor">{r['f_ded']:.6f}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label" style="color:#d4b896;font-weight:600">= CR Final</span>
                    <span class="desglose-valor" style="color:#d4b896">{r['cr_final']:.6f} ‰</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

        with col_b:
            st.markdown("**Construcción de la Prima de Tarifa**")
            cargos_pct = (gasto_admin + gasto_adq + margen_ut) * 100
            st.markdown(f"""
            <div class="card">
                <div class="desglose-row">
                    <span class="desglose-label">Prima de Riesgo</span>
                    <span class="desglose-valor">{fmt_moneda(r['prima_riesgo'])}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">Gastos de Administración ({gasto_admin*100:.0f}%)</span>
                    <span class="desglose-valor">{fmt_moneda(r['g_admin'])}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">Gastos de Adquisición ({gasto_adq*100:.0f}%)</span>
                    <span class="desglose-valor">{fmt_moneda(r['g_adq'])}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">Margen de Utilidad ({margen_ut*100:.0f}%)</span>
                    <span class="desglose-valor">{fmt_moneda(r['g_ut'])}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">Prima de Tarifa (÷{1-(gasto_admin+gasto_adq+margen_ut):.2f})</span>
                    <span class="desglose-valor">{fmt_moneda(r['prima_tarifa'])}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label">Recargo Financiero ({r['recargo_pct']:.2f}%)</span>
                    <span class="desglose-valor">{fmt_moneda(r['prima_con_recargo'] - r['prima_tarifa'])}</span>
                </div>
                <div class="desglose-row">
                    <span class="desglose-label" style="color:#d4b896;font-weight:600">= Prima Total con Recargo</span>
                    <span class="desglose-valor" style="color:#d4b896">{fmt_moneda(r['prima_con_recargo'])}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

        # ── Nota técnica ──
        st.markdown(f"""
        <div class="alerta">
            <strong>Supuestos aplicados:</strong> Gastos totales {cargos_pct:.0f}% · 
            Tasa financiera {tasa_anual*100:.1f}% anual · 
            Deducible fijo ${deducible:,.0f} · 
            Pago {forma_pago.lower()} ({r['m']} {'pago' if r['m']==1 else 'pagos'}) · 
            Factor financiero f = {r['ff']:.6f}<br>
            <em>Esta cotización es de carácter informativo y está basada en estadística SESA 2020-2024.</em>
        </div>
        """, unsafe_allow_html=True)

else:
    # Estado vacío
    st.markdown("""
    <div style="text-align:center; padding: 5rem 2rem; color: #4a5568;">
        <div style="font-size: 3rem; margin-bottom: 1rem;">🏠</div>
        <div style="font-family:'DM Serif Display',serif; font-size:1.4rem; color:#7a8599; margin-bottom:0.5rem;">
            Configura los parámetros en el panel izquierdo
        </div>
        <div style="font-size:0.85rem; color:#4a5568;">
            Selecciona estado · tipo de bien · cobertura · suma asegurada · deducible · forma de pago<br>
            y presiona <strong>Calcular Cotización</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)