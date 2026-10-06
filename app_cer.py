import os
import calendar
from datetime import datetime, date, timedelta
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# =========================================================
# CONFIGURACIÓN DE PÁGINA
# =========================================================
st.set_page_config(page_title="Dashboard Inflación BE (CER)", layout="wide")
st.title("Dashboard: Inflación Implícita Break-Even (Tasa Fija vs CER)")

# =========================================================
# 1. CARGA Y PREPARACIÓN DE DATOS
# =========================================================
@st.cache_data
def cargar_datos_cer():
    archivos = ["Inputs CER_2.xlsx", "Inputs CER.xlsx", "inputs cer.xlsx"]
    file_path = next((a for a in archivos if os.path.exists(a)), None)
            
    if file_path is None:
        st.error("Error: No se encontró el archivo 'Inputs CER_2.xlsx' ni 'Inputs CER.xlsx' en el directorio.")
        st.stop()
        
    xls = pd.ExcelFile(file_path)
    df_lec = pd.read_excel(xls, sheet_name='Lecaps')
    df_cer = pd.read_excel(xls, sheet_name='LECER')
    df_precios = pd.read_excel(xls, sheet_name='Precios')
    df_indice = pd.read_excel(xls, sheet_name='CER')
    df_feriados = pd.read_excel(xls, sheet_name='Feriados')

    # Limpieza de nombres de columnas
    df_lec.columns = df_lec.columns.str.strip()
    df_cer.columns = df_cer.columns.str.strip()
    df_precios.columns = df_precios.columns.str.strip()
    df_indice.columns = df_indice.columns.str.strip()

    # Carga de la hoja Tamar (si existe)
    df_tamar = pd.DataFrame()
    hojas_tamar = [h for h in xls.sheet_names if 'TAMAR' in h.upper()]
    if hojas_tamar:
        df_tamar = pd.read_excel(xls, sheet_name=hojas_tamar[0])
        df_tamar.columns = df_tamar.columns.str.strip()
        col_f = next((c for c in df_tamar.columns if 'FECHA' in c.upper()), df_tamar.columns[0])
        col_v = next((c for c in df_tamar.columns if c != col_f), df_tamar.columns[-1])
        df_tamar['Fecha_dt'] = pd.to_datetime(df_tamar[col_f]).dt.date
        df_tamar['Tamar_val'] = pd.to_numeric(df_tamar[col_v], errors='coerce')
        df_tamar['Tamar_pct'] = df_tamar['Tamar_val'].apply(lambda v: v * 100 if pd.notnull(v) and v < 1.0 else v)
        df_tamar = df_tamar.sort_values('Fecha_dt').reset_index(drop=True)

    if 'Tamar' in df_indice.columns and 'CER' not in df_indice.columns:
        df_indice.rename(columns={'Tamar': 'CER'}, inplace=True)
        
    feriados = set(pd.to_datetime(df_feriados["Feriados"]).dt.date)
    df_indice['Fecha'] = pd.to_datetime(df_indice['Fecha']).dt.date
    df_indice = df_indice.sort_values('Fecha').reset_index(drop=True)
    df_precios['Fecha_dt'] = pd.to_datetime(df_precios['Fecha']).dt.date

    df_lec['Vencimiento_dt'] = pd.to_datetime(df_lec['Vencimiento']).dt.date
    df_cer['Vencimiento_dt'] = pd.to_datetime(df_cer['Vencimiento']).dt.date

    pares = []
    for _, row_lec in df_lec.iterrows():
        vto = row_lec['Vencimiento_dt']
        matching_cer = df_cer[df_cer['Vencimiento_dt'] == vto]
        if not matching_cer.empty:
            for _, row_cer in matching_cer.iterrows():
                pares.append({
                    'label': f"{row_lec['Ticker']} vs {row_cer['Ticker']} (Vto: {vto.strftime('%d/%m/%Y')})",
                    'ticker_lec': str(row_lec['Ticker']).strip().upper(),
                    'ticker_cer': str(row_cer['Ticker']).strip().upper(),
                    'vencimiento': vto
                })

    return df_lec, df_cer, df_precios, df_indice, df_tamar, feriados, pares

df_lec, df_cer, df_precios, df_indice, df_tamar, feriados, pares = cargar_datos_cer()

if not pares:
    st.error("Error: No se encontraron pares de Lecap y LECER con la misma fecha de vencimiento.")
    st.stop()

# =========================================================
# 2. FUNCIONES DE CALENDARIO Y CÁLCULO FINANCIAL MOTOR
# =========================================================
def es_habil(f, feriados):
    return f.weekday() < 5 and f not in feriados

def proximo_habil(f, feriados):
    actual = f
    while not es_habil(actual, feriados):
        actual += timedelta(days=1)
    return actual

def calcular_fecha_liq(f_op, plazo_t, feriados):
    actual = f_op
    sumados = 0
    while sumados < plazo_t:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            sumados += 1
    return proximo_habil(actual, feriados)

def restar_dias_habiles(f, dias_a_restar, feriados):
    actual = f
    restados = 0
    while restados < dias_a_restar:
        actual -= timedelta(days=1)
        if es_habil(actual, feriados):
            restados += 1
    return actual

def dias360_excel(f_inicio, f_fin):
    d1, m1, y1 = f_inicio.day, f_inicio.month, f_inicio.year
    d2, m2, y2 = f_fin.day, f_fin.month, f_fin.year
    if d1 == 31: d1 = 30
    if d2 == 31 and d1 >= 30: d2 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)

def calcular_meses_cer_exactos(f_inicio, f_fin):
    actual = f_inicio
    total_meses = 0.0
    while actual < f_fin:
        dias_en_mes = calendar.monthrange(actual.year, actual.month)[1]
        if actual.year == f_fin.year and actual.month == f_fin.month:
            dias_tramo = (f_fin - actual).days
            total_meses += dias_tramo / dias_en_mes
            break
        else:
            sig_mes = date(actual.year + (1 if actual.month == 12 else 0), 1 if actual.month == 12 else actual.month + 1, 1)
            dias_tramo = (sig_mes - actual).days
            total_meses += dias_tramo / dias_en_mes
            actual = sig_mes
    return total_meses

def calcular_be_cer(ticker_lec, ticker_cer, p_lec, p_cer, f_op, df_lec, df_cer, df_indice, feriados, plazo_t=1):
    row_lec = df_lec[df_lec['Ticker'].str.strip().str.upper() == ticker_lec].iloc[0]
    row_cer = df_cer[df_cer['Ticker'].str.strip().str.upper() == ticker_cer].iloc[0]
    
    f_liq = calcular_fecha_liq(f_op, plazo_t, feriados)
    f_vto_lec = pd.to_datetime(row_lec['Vencimiento']).date()
    f_em_lec = pd.to_datetime(row_lec['Emisión']).date()
    tem_lec = row_lec['Tasa']
    d360_lec = dias360_excel(f_em_lec, f_vto_lec)
    
    vpv_lec = 100.0 * ((1 + tem_lec) ** (d360_lec / 30.0))
    f_cobro_lec = proximo_habil(f_vto_lec, feriados)
    dias_cartera_lec = (f_cobro_lec - f_liq).days
    
    if dias_cartera_lec <= 0 or p_lec <= 0: return None
        
    rend_lec = (vpv_lec / p_lec) - 1
    tir_lec = ((1 + rend_lec) ** (365 / dias_cartera_lec)) - 1
    
    f_vto_cer = pd.to_datetime(row_cer['Vencimiento']).date()
    f_cobro_cer = proximo_habil(f_vto_cer, feriados)
    dias_cartera_cer = (f_cobro_cer - f_liq).days
    
    if dias_cartera_cer <= 0 or p_cer <= 0: return None
        
    pesos_requeridos_cer = p_cer * ((1 + tir_lec) ** (dias_cartera_cer / 365))
    
    cer_inicial_emision = row_cer['CER inicial']
    cupon_cer = row_cer['Tasa']
    vpv_real = 100.0 * (1 + cupon_cer)
    cer_vto_req = (pesos_requeridos_cer / vpv_real) * cer_inicial_emision
    
    f_ref_vto = restar_dias_habiles(f_cobro_cer, 10, feriados)
    f_ref_hoy = restar_dias_habiles(f_liq, 10, feriados)
    
    df_hist_hoy = df_indice[df_indice['Fecha'] <= f_ref_hoy]
    cer_hoy_ref = df_hist_hoy['CER'].iloc[-1] if not df_hist_hoy.empty else cer_inicial_emision
    
    # Cálculo de Tasa Real del CER (TIR Real Anualizada del LECER)
    factor_real = (vpv_real * cer_hoy_ref) / (p_cer * cer_inicial_emision)
    tir_real_cer = (factor_real ** (365.0 / dias_cartera_cer)) - 1.0 if factor_real > 0 else 0.0

    inflacion_total_be = (cer_vto_req / cer_hoy_ref) - 1
    meses_exactos = calcular_meses_cer_exactos(f_ref_hoy, f_ref_vto)
    
    if meses_exactos > 0:
        tem_be = ((1 + inflacion_total_be) ** (1.0 / meses_exactos)) - 1
        tea_be = ((1 + inflacion_total_be) ** (12.0 / meses_exactos)) - 1
    else:
        tem_be, tea_be = 0.0, 0.0
        
    return {
        'fecha_op': f_op, 'vpv_lec': vpv_lec, 'tir_lec': tir_lec,
        'pesos_req_cer': pesos_requeridos_cer, 'cer_hoy_ref': cer_hoy_ref,
        'cer_vto_req': cer_vto_req, 'f_ref_hoy': f_ref_hoy, 'f_ref_vto': f_ref_vto,
        'meses_exactos': meses_exactos, 'infla_total': inflacion_total_be,
        'tem_be': tem_be, 'tea_be': tea_be, 'tir_real_cer': tir_real_cer
    }

# =========================================================
# 3. INTERFAZ Y BARRA LATERAL
# =========================================================
ultima_fila = df_precios.iloc[-1]
fecha_sim = ultima_fila['Fecha_dt']

st.sidebar.header("Parámetros de Simulación")
st.sidebar.write(f"**Fecha del último dato:** {fecha_sim.strftime('%d/%m/%Y')}")

opciones_labels = [p['label'] for p in pares]

# Índice por defecto para S30N6 vs X30N6 si existe
default_idx = next(
    (i for i, p in enumerate(pares) if p['ticker_lec'] == 'S30N6' and p['ticker_cer'] == 'X30N6'), 
    0
)

par_seleccionado_label = st.sidebar.selectbox(
    "1. Seleccione el par a comparar (Mismo Vencimiento):", 
    opciones_labels, 
    index=default_idx
)

par_info = next(p for p in pares if p['label'] == par_seleccionado_label)
lec_ticker = par_info['ticker_lec']
cer_ticker = par_info['ticker_cer']

val_l = ultima_fila[lec_ticker] if lec_ticker in df_precios.columns else 100.0
val_c = ultima_fila[cer_ticker] if cer_ticker in df_precios.columns else 100.0

def_p_l = float(val_l * 100 if pd.notnull(val_l) and val_l < 10 else val_l)
def_p_c = float(val_c * 100 if pd.notnull(val_c) and val_c < 10 else val_c)

st.sidebar.markdown("---")
precio_lec_sim = st.sidebar.number_input(f"Precio {lec_ticker} (Tasa Fija):", value=def_p_l, step=0.10, format="%.2f")
precio_cer_sim = st.sidebar.number_input(f"Precio {cer_ticker} (CER):", value=def_p_c, step=0.10, format="%.2f")

st.sidebar.markdown("---")
st.sidebar.subheader("Rango Histórico a Graficar")
fecha_min_gen = df_precios['Fecha_dt'].min()
fecha_max_gen = df_precios['Fecha_dt'].max()

rango_fechas = st.sidebar.date_input("Seleccione el período:",
    value=(fecha_min_gen, fecha_max_gen),
    min_value=fecha_min_gen, max_value=fecha_max_gen)

if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
    f_desde, f_hasta = rango_fechas
else:
    f_desde, f_hasta = fecha_min_gen, fecha_max_gen

# =========================================================
# 4. EJECUCIÓN Y RENDERIZADO DE MÉTRICAS
# =========================================================
res_sim = calcular_be_cer(lec_ticker, cer_ticker, precio_lec_sim, precio_cer_sim, fecha_sim, 
                          df_lec, df_cer, df_indice, feriados)

if res_sim:
    st.subheader(f"Resultados de Simulación: {lec_ticker} vs {cer_ticker}")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(label=f"TEA Fija ({lec_ticker})", value=f"{res_sim['tir_lec']*100:.2f}%")
    c2.metric(label=f"Tasa Real CER ({cer_ticker})", value=f"{res_sim['tir_real_cer']*100:.2f}%")
    c3.metric(label=f"Pesos Target ({cer_ticker})", value=f"${res_sim['pesos_req_cer']:.2f}")
    c4.metric(label="TEM BE (Mensual)", value=f"{res_sim['tem_be']*100:.2f}%")
    c5.metric(label="TEA BE (TEM Anualizada)", value=f"{res_sim['tea_be']*100:.2f}%")

# Serie Histórica
df_precios_filt = df_precios[(df_precios['Fecha_dt'] >= f_desde) & (df_precios['Fecha_dt'] <= f_hasta)]
historico = []

for _, row in df_precios_filt.iterrows():
    f_o = row['Fecha_dt']
    p_l_raw = row[lec_ticker] if lec_ticker in row and pd.notnull(row[lec_ticker]) else None
    p_c_raw = row[cer_ticker] if cer_ticker in row and pd.notnull(row[cer_ticker]) else None
    
    if p_l_raw is not None and p_c_raw is not None:
        p_l = p_l_raw * 100 if p_l_raw < 10 else p_l_raw
        p_c = p_c_raw * 100 if p_c_raw < 10 else p_c_raw
        
        r = calcular_be_cer(lec_ticker, cer_ticker, p_l, p_c, f_o, df_lec, df_cer, df_indice, feriados)
        if r:
            if not df_tamar.empty:
                tamar_sub = df_tamar[df_tamar['Fecha_dt'] <= f_o]
                r['tamar'] = tamar_sub['Tamar_pct'].iloc[-1] if not tamar_sub.empty else np.nan
            else:
                r['tamar'] = np.nan
            historico.append(r)

df_hist = pd.DataFrame(historico)

# =========================================================
# 5. GRÁFICOS INTERACTIVOS (PLOTLY)
# =========================================================
st.markdown("---")
st.subheader("Evolución Histórica de Métricas")

if not df_hist.empty:
    df_hist['tem_be_pct'] = df_hist['tem_be'] * 100
    df_hist['tea_be_pct'] = df_hist['tea_be'] * 100
    df_hist['tir_real_cer_pct'] = df_hist['tir_real_cer'] * 100

    # ---------------------------------------------------------
    # GRÁFICO 1: TEM BE (Inflación Mensual Implícita)
    # ---------------------------------------------------------
    fig1 = go.Figure()
    fig1.add_trace(go.Scatter(
        x=df_hist['fecha_op'], y=df_hist['tem_be_pct'],
        name="TEM BE (%)", line=dict(color='#d62728', width=2.5),
        hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEM BE:</b> %{y:.2f}%<extra></extra>"
    ))
    fig1.update_layout(
        title=f"<b>1. Inflación Mensual Implícita (TEM BE): {lec_ticker} vs {cer_ticker}</b>",
        xaxis_title="Fecha", yaxis_title="TEM (%)",
        template="plotly_white", hovermode="x unified"
    )
    st.plotly_chart(fig1, use_container_width=True)

    # ---------------------------------------------------------
    # GRÁFICO 2: TEA BE vs TASA REAL CER (Doble Eje)
    # ---------------------------------------------------------
    fig2 = make_subplots(specs=[[{"secondary_y": True}]])
    fig2.add_trace(
        go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tea_be_pct'],
            name="TEA BE (%)", line=dict(color='#1f77b4', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEA BE:</b> %{y:.2f}%<extra></extra>"
        ),
        secondary_y=False
    )
    fig2.add_trace(
        go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tir_real_cer_pct'],
            name=f"Tasa Real CER ({cer_ticker})", line=dict(color='#2ca02c', width=2.5, dash='solid'),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Tasa Real CER:</b> %{y:.2f}%<extra></extra>"
        ),
        secondary_y=True
    )
    fig2.update_layout(
        title=f"<b>2. Inflación Anualizada Implícita (TEA BE) vs Tasa Real CER ({cer_ticker})</b>",
        template="plotly_white", hovermode="x unified"
    )
    fig2.update_xaxes(title_text="Fecha")
    fig2.update_yaxes(title_text="<b>TEA BE (%)</b>", secondary_y=False, title_font_color='#1f77b4')
    fig2.update_yaxes(title_text="<b>Tasa Real CER (%)</b>", secondary_y=True, title_font_color='#2ca02c')
    st.plotly_chart(fig2, use_container_width=True)

    # ---------------------------------------------------------
    # GRÁFICO 3: TEA BE vs TASA TAMAR (Doble Eje)
    # ---------------------------------------------------------
    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
    fig3.add_trace(
        go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tea_be_pct'],
            name="TEA BE (%)", line=dict(color='#1f77b4', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEA BE:</b> %{y:.2f}%<extra></extra>"
        ),
        secondary_y=False
    )
    fig3.add_trace(
        go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tamar'],
            name="Tasa TAMAR (%)", line=dict(color='#ff7f0e', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR:</b> %{y:.2f}%<extra></extra>"
        ),
        secondary_y=True
    )
    fig3.update_layout(
        title=f"<b>3. Inflación Anualizada Implícita (TEA BE) vs Tasa TAMAR</b>",
        template="plotly_white", hovermode="x unified"
    )
    fig3.update_xaxes(title_text="Fecha")
    fig3.update_yaxes(title_text="<b>TEA BE (%)</b>", secondary_y=False, title_font_color='#1f77b4')
    fig3.update_yaxes(title_text="<b>Tasa TAMAR (%)</b>", secondary_y=True, title_font_color='#ff7f0e')
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.warning("No hay suficientes datos históricos para el rango seleccionado.")