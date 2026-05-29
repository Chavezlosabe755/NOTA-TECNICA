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

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp { background-color: #0f1117; color: #e8e8e8; }

[data-testid="stSidebar"] {
    background-color: #161b27;
    border-right: 1px solid #2a3142;
}

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
.card-sub { font-size: 0.82rem; color: #7a8599; margin-top: 0.2rem; }

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

/* Cabecera de tabla de coberturas */
.cob-header {
    display: flex;
    justify-content: space-between;
    font-size: 0.72rem;
    color: #7a8599;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    padding: 0.4rem 0;
    border-bottom: 1px solid #2a3142;
    margin-bottom: 0.2rem;
}

.alerta {
    background: #1e1a2e;
    border-left: 3px solid #d4b896;
    border-radius: 6px;
    padding: 0.8rem 1rem;
    font-size: 0.85rem;
    color: #a0aabb;
    margin-top: 1rem;
}

.seccion-titulo {
    font-size: 0.72rem;
    color: #7a8599;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin: 1.4rem 0 0.6rem 0;
    border-bottom: 1px solid #2a3142;
    padding-bottom: 0.4rem;
}

.total-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.6rem 0;
    font-size: 1rem;
    border-top: 2px solid #d4b896;
    margin-top: 0.4rem;
}

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
    cr_global    = pd.read_excel(xls, "CR_Global")
    factor_geo   = pd.read_excel(xls, "Factor_Geografico")
    factor_ded   = pd.read_excel(xls, "Factor_Deducible")
    prima_tarifa = pd.read_excel(xls, "Prima_Tarifa")
    return cr_global, factor_geo, factor_ded, prima_tarifa


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt_moneda(valor):
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return "—"
    return f"${valor:,.2f}"

def factor_financiero(j, m):
    denominador = sum((1 + j / m) ** (-k) for k in range(m))
    return m / denominador

def calcular_una_cobertura(
    cr_global_df, factor_geo_df, factor_ded_df,
    cobertura, tipo_bien, estado,
    suma_asegurada, deducible,
    gasto_admin, gasto_adq, margen_ut,
    tasa_anual, m, ff,
):
    """Calcula prima para UNA cobertura. Devuelve dict o None si no hay datos."""

    # 1. CR Global
    mask_g = (
        (cr_global_df["COBERTURA"]    == cobertura) &
        (cr_global_df["TIPO DE BIEN"] == tipo_bien)
    )
    fila_g = cr_global_df[mask_g]
    if fila_g.empty:
        return None
    cr_global = fila_g["CR_GLOBAL"].values[0]
    if pd.isna(cr_global) or cr_global == 0:
        return None

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

    # 3. Factor deducible
    fila_ded = factor_ded_df[factor_ded_df["DEDUCIBLE"] == deducible]
    f_ded = fila_ded["f"].values[0] if not fila_ded.empty else 1.0
    if pd.isna(f_ded):
        f_ded = 1.0

    # 4. Cuota y primas
    cr_final     = cr_global * f_geo * f_ded
    prima_riesgo = suma_asegurada * cr_final / 1000
    cargos       = gasto_admin + gasto_adq + margen_ut
    prima_tarifa = prima_riesgo / (1 - cargos)
    prima_recargo = prima_tarifa * ff if m > 1 else prima_tarifa
    recibo        = prima_recargo / m

    return {
        "cobertura"    : cobertura,
        "cr_global"    : cr_global,
        "f_geo"        : f_geo,
        "f_ded"        : f_ded,
        "cr_final"     : cr_final,
        "prima_riesgo" : prima_riesgo,
        "prima_tarifa" : prima_tarifa,
        "prima_recargo": prima_recargo,
        "recibo"       : recibo,
        "g_admin"      : prima_tarifa * gasto_admin,
        "g_adq"        : prima_tarifa * gasto_adq,
        "g_ut"         : prima_tarifa * margen_ut,
    }


# ─────────────────────────────────────────────
# ENCABEZADO
# ─────────────────────────────────────────────
_, col_titulo = st.columns([1, 5])
with col_titulo:
    st.markdown('<div class="titulo-principal">Cotizador<br>Casa Habitación</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitulo">Bouclier Seguros de Daños · Ramo Incendio</div>', unsafe_allow_html=True)

st.divider()

# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────
if not os.path.exists(EXCEL_PATH):
    st.error(
        f"⚠️ No se encontró **{EXCEL_PATH}** en el directorio actual.\n\n"
        "Corre primero el notebook y coloca el Excel en la misma carpeta que este script."
    )
    st.stop()

try:
    cr_global_df, factor_geo_df, factor_ded_df, prima_tarifa_df = cargar_datos(EXCEL_PATH)
except Exception as e:
    st.error(f"Error al leer el Excel: {e}")
    st.stop()

coberturas_disponibles = sorted(cr_global_df["COBERTURA"].dropna().unique().tolist())
tipos_bien_disponibles = sorted(cr_global_df["TIPO DE BIEN"].dropna().unique().tolist())
estados_disponibles    = sorted(factor_geo_df["ENTIDAD"].dropna().unique().tolist())
deducibles_disponibles = sorted(factor_ded_df["DEDUCIBLE"].dropna().unique().tolist())


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="seccion-titulo">Datos del Riesgo</div>', unsafe_allow_html=True)

    estado = st.selectbox("Entidad Federativa", estados_disponibles)

    tipo_bien = st.selectbox(
        "Tipo de Bien",
        tipos_bien_disponibles,
        help="Contenidos, Edificio o Contenidos y Edificio"
    )

    # ── MULTISELECT de coberturas ──
    coberturas_sel = st.multiselect(
        "Coberturas",
        coberturas_disponibles,
        default=[coberturas_disponibles[0]] if coberturas_disponibles else [],
        help="Puedes seleccionar una o varias coberturas; la prima total es la suma de cada una."
    )

    suma_asegurada = st.number_input(
        "Suma Asegurada ($)",
        min_value=10_000.0,
        max_value=500_000_000.0,
        value=1_000_000.0,
        step=10_000.0,
        format="%.2f",
    )

    deducible = st.selectbox(
        "Deducible ($)",
        deducibles_disponibles,
        index=0,
        format_func=lambda x: f"${x:,.0f}",
    )

    st.markdown('<div class="seccion-titulo">Forma de Pago</div>', unsafe_allow_html=True)

    forma_pago = st.radio(
        "Periodicidad",
        ["Anual", "Semestral", "Trimestral", "Mensual"],
    )

    # Supuestos técnicos fijos (no editables por el usuario)
    gasto_admin = 0.20
    gasto_adq   = 0.05
    margen_ut   = 0.06
    tasa_anual  = 0.08

    cotizar = st.button("Calcular Cotización", use_container_width=True, type="primary")


# ─────────────────────────────────────────────
# ÁREA PRINCIPAL
# ─────────────────────────────────────────────
if cotizar:

    if not coberturas_sel:
        st.warning("⚠️ Selecciona al menos una cobertura.")
        st.stop()

    cargos = gasto_admin + gasto_adq + margen_ut
    if cargos >= 1:
        st.warning("⚠️ Los gastos totales no pueden ser ≥ 100%.")
        st.stop()

    pagos_map = {"Anual": 1, "Semestral": 2, "Trimestral": 4, "Mensual": 12}
    m  = pagos_map[forma_pago]
    ff = factor_financiero(tasa_anual, m) if m > 1 else 1.0

    # Calcular cada cobertura seleccionada
    resultados = []
    sin_datos  = []

    for cob in coberturas_sel:
        r = calcular_una_cobertura(
            cr_global_df, factor_geo_df, factor_ded_df,
            cob, tipo_bien, estado,
            suma_asegurada, deducible,
            gasto_admin, gasto_adq, margen_ut,
            tasa_anual, m, ff,
        )
        if r:
            resultados.append(r)
        else:
            sin_datos.append(cob)

    if sin_datos:
        st.warning(f"⚠️ Sin datos para: {', '.join(sin_datos)}. Se excluyen del cálculo.")

    if not resultados:
        st.error("No hay datos disponibles para ninguna de las coberturas seleccionadas.")
        st.stop()

    # Totales
    total_prima_riesgo  = sum(r["prima_riesgo"]  for r in resultados)
    total_prima_tarifa  = sum(r["prima_tarifa"]  for r in resultados)
    total_prima_recargo = sum(r["prima_recargo"] for r in resultados)
    total_recibo        = sum(r["recibo"]        for r in resultados)

    # ── Resumen del riesgo ──
    st.markdown('<div class="seccion-titulo">Resumen del Riesgo</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Estado", estado)
    c2.metric("Tipo de Bien", tipo_bien)
    c3.metric("Coberturas", str(len(resultados)))
    c4.metric("Suma Asegurada", fmt_moneda(suma_asegurada))

    # ── Tarjetas resumen ──
    st.markdown('<div class="seccion-titulo">Resultado de la Cotización</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(f"""
        <div class="card">
            <div class="card-label">Prima Total Anual</div>
            <div class="card-value">{fmt_moneda(total_prima_recargo)}</div>
            <div class="card-sub">{len(resultados)} cobertura{'s' if len(resultados)>1 else ''} · incluye recargo financiero</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        label_pago = {1:"Pago Único", 2:"Recibo Semestral", 4:"Recibo Trimestral", 12:"Recibo Mensual"}[m]
        st.markdown(f"""
        <div class="card">
            <div class="card-label">{label_pago}</div>
            <div class="card-value">{fmt_moneda(total_recibo)}</div>
            <div class="card-sub">{m} pago{'s' if m>1 else ''} de igual valor</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        recargo_monto = total_prima_recargo - total_prima_tarifa
        recargo_pct   = (ff - 1) * 100 if m > 1 else 0.0
        st.markdown(f"""
        <div class="card">
            <div class="card-label">Prima de Riesgo Total</div>
            <div class="card-value">{fmt_moneda(total_prima_riesgo)}</div>
            <div class="card-sub">Recargo financiero: {recargo_pct:.2f}%</div>
        </div>
        """, unsafe_allow_html=True)

    # ── Desglose por cobertura ──
    st.markdown('<div class="seccion-titulo">Desglose por Cobertura</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        <div class="cob-header">
            <span style="flex:2">Cobertura</span>
            <span style="flex:1;text-align:right">CR Final ‰</span>
            <span style="flex:1;text-align:right">f. Geo</span>
            <span style="flex:1;text-align:right">f. Deducible</span>
            <span style="flex:1;text-align:right">Prima Riesgo</span>
            <span style="flex:1;text-align:right">Prima Tarifa</span>
            <span style="flex:1;text-align:right">Prima c/Recargo</span>
            <span style="flex:1;text-align:right">Recibo</span>
        </div>
    """, unsafe_allow_html=True)

    filas_html = ""
    for r in resultados:
        filas_html += f"""
        <div class="desglose-row">
            <span class="desglose-label" style="flex:2">{r['cobertura']}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{r['cr_final']:.6f}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{r['f_geo']:.4f}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{r['f_ded']:.4f}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{fmt_moneda(r['prima_riesgo'])}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{fmt_moneda(r['prima_tarifa'])}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{fmt_moneda(r['prima_recargo'])}</span>
            <span class="desglose-valor" style="flex:1;text-align:right">{fmt_moneda(r['recibo'])}</span>
        </div>
        """

    # Fila de totales
    filas_html += f"""
    <div class="total-row">
        <span style="flex:2;font-weight:600;color:#d4b896">TOTAL</span>
        <span style="flex:1"></span>
        <span style="flex:1"></span>
        <span style="flex:1"></span>
        <span style="flex:1;text-align:right;font-weight:600;color:#d4b896">{fmt_moneda(total_prima_riesgo)}</span>
        <span style="flex:1;text-align:right;font-weight:600;color:#d4b896">{fmt_moneda(total_prima_tarifa)}</span>
        <span style="flex:1;text-align:right;font-weight:600;color:#d4b896">{fmt_moneda(total_prima_recargo)}</span>
        <span style="flex:1;text-align:right;font-weight:600;color:#d4b896">{fmt_moneda(total_recibo)}</span>
    </div>
    """

    st.markdown(filas_html + "</div>", unsafe_allow_html=True)

    # ── Desglose de gastos (totales) ──
    st.markdown('<div class="seccion-titulo">Construcción de la Prima Total</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown(f"""
        <div class="card">
            <div class="desglose-row">
                <span class="desglose-label">Prima de Riesgo (suma coberturas)</span>
                <span class="desglose-valor">{fmt_moneda(total_prima_riesgo)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label">Gastos de Administración ({gasto_admin*100:.0f}%)</span>
                <span class="desglose-valor">{fmt_moneda(total_prima_tarifa * gasto_admin)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label">Gastos de Adquisición ({gasto_adq*100:.0f}%)</span>
                <span class="desglose-valor">{fmt_moneda(total_prima_tarifa * gasto_adq)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label">Margen de Utilidad ({margen_ut*100:.0f}%)</span>
                <span class="desglose-valor">{fmt_moneda(total_prima_tarifa * margen_ut)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label" style="color:#d4b896;font-weight:600">= Prima de Tarifa</span>
                <span class="desglose-valor" style="color:#d4b896">{fmt_moneda(total_prima_tarifa)}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_b:
        st.markdown(f"""
        <div class="card">
            <div class="desglose-row">
                <span class="desglose-label">Prima de Tarifa</span>
                <span class="desglose-valor">{fmt_moneda(total_prima_tarifa)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label">Factor Financiero (j={tasa_anual*100:.1f}%, m={m})</span>
                <span class="desglose-valor">{ff:.6f}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label">Recargo Financiero ({recargo_pct:.2f}%)</span>
                <span class="desglose-valor">{fmt_moneda(recargo_monto)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label" style="color:#d4b896;font-weight:600">= Prima Total con Recargo</span>
                <span class="desglose-valor" style="color:#d4b896">{fmt_moneda(total_prima_recargo)}</span>
            </div>
            <div class="desglose-row">
                <span class="desglose-label" style="color:#d4b896;font-weight:600">= {label_pago}</span>
                <span class="desglose-valor" style="color:#d4b896">{fmt_moneda(total_recibo)}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Nota de supuestos ──
    cobs_txt = " · ".join([r["cobertura"] for r in resultados])
    st.markdown(f"""
    <div class="alerta">
        <strong>Coberturas:</strong> {cobs_txt}<br>
        <strong>Supuestos:</strong> Gastos totales {cargos*100:.0f}% · 
        Tasa financiera {tasa_anual*100:.1f}% anual · 
        Deducible ${deducible:,.0f} · 
        Pago {forma_pago.lower()} ({m} {'pago' if m==1 else 'pagos'}) · 
        f = {ff:.6f}<br>
        <em>Cotización informativa basada en estadística SESA 2020-2024.</em>
    </div>
    """, unsafe_allow_html=True)

else:
    st.markdown("""
    <div style="text-align:center; padding: 5rem 2rem; color: #4a5568;">
        <div style="font-size: 3rem; margin-bottom: 1rem;">🏠</div>
        <div style="font-family:'DM Serif Display',serif; font-size:1.4rem; color:#7a8599; margin-bottom:0.5rem;">
            Configura los parámetros en el panel izquierdo
        </div>
        <div style="font-size:0.85rem; color:#4a5568;">
            Estado · tipo de bien · coberturas · suma asegurada · deducible · forma de pago<br>
            y presiona <strong>Calcular Cotización</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)