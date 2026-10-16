import os
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import requests

# =========================================================
# CONFIGURACIÓN DE PÁGINA
# =========================================================
st.set_page_config(page_title="Dashboard Inflación BE (CER)", layout="wide")
st.title("Dashboard: Inflación Implícita Break-Even (Tasa Fija vs CER)")

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# Overrides manuales de VPV (idealmente pasarlos a la columna 'VPV' de la hoja Lecaps)
VPV_OVERRIDE = {'S29E7': 111.6660}

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

    for d in (df_lec, df_cer, df_precios, df_indice):
        d.columns = d.columns.str.strip()

    # Hoja Tamar (opcional)
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

    # Limpieza de la curva CER: descartar filas constantes repetidas al final
    diffs = df_indice['CER'].diff().abs()
    if not diffs.empty and (diffs > 0).any():
        last_changing_idx = df_indice[diffs > 0].index[-1]
        df_indice_clean = df_indice.loc[:last_changing_idx].copy()
    else:
        df_indice_clean = df_indice.copy()

    df_precios['Fecha_dt'] = pd.to_datetime(df_precios['Fecha']).dt.date
    df_precios = df_precios.sort_values('Fecha_dt').reset_index(drop=True)

    df_lec['Vencimiento_dt'] = pd.to_datetime(df_lec['Vencimiento']).dt.date
    df_cer['Vencimiento_dt'] = pd.to_datetime(df_cer['Vencimiento']).dt.date

    pares = []
    for _, row_lec in df_lec.iterrows():
        vto = row_lec['Vencimiento_dt']
        matching_cer = df_cer[df_cer['Vencimiento_dt'] == vto]
        for _, row_cer in matching_cer.iterrows():
            pares.append({
                'label': f"{row_lec['Ticker']} vs {row_cer['Ticker']} (Vto: {vto.strftime('%d/%m/%Y')})",
                'ticker_lec': str(row_lec['Ticker']).strip().upper(),
                'ticker_cer': str(row_cer['Ticker']).strip().upper(),
                'vencimiento': vto
            })

    return df_lec, df_cer, df_precios, df_indice_clean, df_tamar, feriados, pares

df_lec, df_cer, df_precios, df_indice, df_tamar, feriados, pares = cargar_datos_cer()

if not pares:
    st.error("Error: No se encontraron pares de Lecap y LECER con la misma fecha de vencimiento.")
    st.stop()

# =========================================================
# 2. CALENDARIO
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

def _es_ultimo_dia_feb(f):
    return f.month == 2 and (f + timedelta(days=1)).month == 3

def dias360_excel(f_inicio, f_fin):
    """DAYS360 método US (NASD), igual al default de Excel."""
    d1, m1, y1 = f_inicio.day, f_inicio.month, f_inicio.year
    d2, m2, y2 = f_fin.day, f_fin.month, f_fin.year
    if _es_ultimo_dia_feb(f_inicio) and _es_ultimo_dia_feb(f_fin):
        d2 = 30
    if _es_ultimo_dia_feb(f_inicio):
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    if d1 == 31:
        d1 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)

def normalizar_precio(x):
    """Precios pueden venir como 1.05 o 105. Devuelve base 100. NaN-safe."""
    if x is None or pd.isna(x):
        return np.nan
    x = float(x)
    return x * 100 if x < 10 else x

# =========================================================
# 3. FUNCIONES DE CER
# =========================================================
# Los "períodos CER" van del 16 de un mes al 15 del siguiente. Dentro de cada período
# el CER crece a tasa diaria geométrica constante (según el IPC ya publicado).
def id_periodo_cer(f):
    """(año, mes) del día 16 con que arranca el período CER que contiene a f."""
    if f.day >= 16:
        return (f.year, f.month)
    return (f.year - 1, 12) if f.month == 1 else (f.year, f.month - 1)

def fin_periodo_cer(f):
    """Día 15 con que termina el período CER que contiene a f."""
    if f.day <= 15:
        return date(f.year, f.month, 15)
    return date(f.year + 1, 1, 15) if f.month == 12 else date(f.year, f.month + 1, 15)

def calcular_meses_efectivos_bcra(f_inicio, f_fin):
    """N efectivo según Com. "A" 3507 del BCRA (períodos 16 → 15)."""
    if f_inicio >= f_fin:
        return 0.0

    n_efectivo = 0.0
    actual = f_inicio

    while actual < f_fin:
        if actual.day >= 16:
            inicio_periodo = date(actual.year, actual.month, 16)
            fin_periodo = date(actual.year + 1, 1, 15) if actual.month == 12 else date(actual.year, actual.month + 1, 15)
        else:
            fin_periodo = date(actual.year, actual.month, 15)
            inicio_periodo = date(actual.year - 1, 12, 16) if actual.month == 1 else date(actual.year, actual.month - 1, 16)

        dias_totales_periodo = (fin_periodo - inicio_periodo).days + 1
        fin_tramo = min(f_fin, fin_periodo)
        dias_en_tramo = (fin_tramo - actual).days if actual == f_inicio else (fin_tramo - actual).days + 1

        if dias_en_tramo > 0:
            n_efectivo += dias_en_tramo / dias_totales_periodo

        actual = fin_periodo + timedelta(days=1)

    return n_efectivo

def obtener_limite_cer_historico(f_op):
    """Hasta qué fecha el mercado conocía el CER en f_op (IPC se publica ~ día 14-15)."""
    if f_op.day < 15:
        return date(f_op.year, f_op.month, 15)
    return date(f_op.year + 1, 1, 15) if f_op.month == 12 else date(f_op.year, f_op.month + 1, 15)

def cer_conocido(df_indice, f_corte):
    """
    Último CER publicado <= f_corte. Si los datos terminan antes de f_corte pero dentro del
    mismo período CER (p.ej. datos hasta 13/11 y el período cierra el 15/11), completa
    extrapolando con la tasa diaria constante observada en ese período.
    Devuelve (fecha, cer) o (None, None).
    """
    d = df_indice[df_indice['Fecha'] <= f_corte]
    if d.empty:
        return None, None
    f_u, c_u = d['Fecha'].iloc[-1], float(d['CER'].iloc[-1])

    if f_u < f_corte and id_periodo_cer(f_u) == id_periodo_cer(f_corte):
        y, m = id_periodo_cer(f_corte)
        base = d[(d['Fecha'] >= date(y, m, 16)) & (d['Fecha'] < f_u)]
        if not base.empty:
            f_b, c_b = base['Fecha'].iloc[0], float(base['CER'].iloc[0])
            dias = (f_u - f_b).days
            if dias > 0 and c_b > 0:
                g = (c_u / c_b) ** (1.0 / dias)
                c_u = c_u * g ** (f_corte - f_u).days
                f_u = f_corte
    return f_u, c_u

# =========================================================
# 4. MOTOR DE CÁLCULO
# =========================================================
def calcular_be_cer(ticker_lec, ticker_cer, p_lec, p_cer, f_op, df_lec, df_cer, df_indice, feriados,
                    plazo_t=1, modo_historico=True):
    """
    modo_historico=True : el CER conocido en f_op se limita por la regla de publicación del IPC
                          (evita look-ahead en la serie histórica).
    modo_historico=False: simulación "hoy": se usa todo el CER cargado en el Excel; el tramo
                          conocido llega hasta el cierre del período CER de la última fecha cargada.
    """
    if pd.isna(p_lec) or pd.isna(p_cer) or p_lec <= 0 or p_cer <= 0:
        return None

    row_lec = df_lec[df_lec['Ticker'].str.strip().str.upper() == ticker_lec].iloc[0]
    row_cer = df_cer[df_cer['Ticker'].str.strip().str.upper() == ticker_cer].iloc[0]

    f_liq = calcular_fecha_liq(f_op, plazo_t, feriados)

    # --- LECAP ---
    f_vto_lec = pd.to_datetime(row_lec['Vencimiento']).date()
    f_em_lec = pd.to_datetime(row_lec['Emisión']).date()
    tem_lec = row_lec['Tasa']
    d360_lec = dias360_excel(f_em_lec, f_vto_lec)

    if ticker_lec in VPV_OVERRIDE:
        vpv_lec = VPV_OVERRIDE[ticker_lec]
    elif 'VPV' in row_lec.index and pd.notnull(row_lec['VPV']):
        vpv_lec = float(row_lec['VPV'])
    else:
        vpv_lec = 100.0 * ((1 + tem_lec) ** (d360_lec / 30.0))

    f_cobro_lec = proximo_habil(f_vto_lec, feriados)
    dias_cartera_lec = (f_cobro_lec - f_liq).days
    if dias_cartera_lec <= 0:
        return None

    rend_lec = (vpv_lec / p_lec) - 1.0
    tir_lec = ((1.0 + rend_lec) ** (365.0 / dias_cartera_lec)) - 1.0

    # --- LECER ---
    f_vto_cer = pd.to_datetime(row_cer['Vencimiento']).date()
    f_cobro_cer = proximo_habil(f_vto_cer, feriados)
    dias_cartera_cer = (f_cobro_cer - f_liq).days
    if dias_cartera_cer <= 0:
        return None

    pesos_requeridos_cer = p_cer * ((1.0 + tir_lec) ** (dias_cartera_cer / 365.0))

    cer_inicial_emision = float(row_cer['CER inicial'])
    cupon_cer = row_cer['Tasa']
    vpv_real = 100.0 * (1.0 + cupon_cer)
    cer_vto_req = (pesos_requeridos_cer / vpv_real) * cer_inicial_emision

    # CER de referencia al vencimiento (rezago 10 hábiles)
    f_ref_vto = restar_dias_habiles(f_cobro_cer, 10, feriados)

    # CER de referencia hoy y tasa real (con CER constante)
    f_ref_hoy = restar_dias_habiles(f_liq, 10, feriados)
    df_hist_hoy = df_indice[df_indice['Fecha'] <= f_ref_hoy]
    cer_hoy_ref = float(df_hist_hoy['CER'].iloc[-1]) if not df_hist_hoy.empty else cer_inicial_emision
    factor_real = (vpv_real * cer_hoy_ref) / (p_cer * cer_inicial_emision)
    tir_real_cer = (factor_real ** (365.0 / dias_cartera_cer)) - 1.0 if factor_real > 0 else np.nan

    # ---------------------------------------------------------
    # ¿Cuánto del CER ya se conoce?
    # ---------------------------------------------------------
    if modo_historico:
        f_limite = obtener_limite_cer_historico(f_op)
    else:
        f_limite = fin_periodo_cer(df_indice['Fecha'].max())

    f_corte = min(f_ref_vto, f_limite)
    f_cer_ult, cer_ult = cer_conocido(df_indice, f_corte)
    if f_cer_ult is None or cer_ult <= 0:
        return None

    cer_es_conocido = f_ref_vto <= f_limite
    tem_be = tea_be = n_meses = infla_desconocida = tir_cer_fija = np.nan

    if cer_es_conocido:
        # El CER al vencimiento ya está determinado: no hay inflación que "implicar".
        # La LECER es, en la práctica, de tasa fija -> se compara TIR contra TIR.
        if f_cer_ult < f_ref_vto:
            return None  # faltan datos de CER para completar el tramo conocido
        pesos_vto_cer = vpv_real * cer_ult / cer_inicial_emision
        tir_cer_fija = (pesos_vto_cer / p_cer) ** (365.0 / dias_cartera_cer) - 1.0
    else:
        n_meses = calcular_meses_efectivos_bcra(f_cer_ult, f_ref_vto)
        if n_meses <= 0 or cer_vto_req <= 0:
            return None
        infla_desconocida = (cer_vto_req / cer_ult) - 1.0
        tem_be = ((cer_vto_req / cer_ult) ** (1.0 / n_meses)) - 1.0
        tea_be = ((1.0 + tem_be) ** 12.0) - 1.0

    return {
        'fecha_op': f_op, 'vpv_lec': vpv_lec, 'tir_lec': tir_lec,
        'pesos_req_cer': pesos_requeridos_cer, 'cer_hoy_ref': cer_hoy_ref,
        'cer_vto_req': cer_vto_req, 'f_ref_hoy': f_ref_hoy, 'f_ref_vto': f_ref_vto,
        'infla_total': (cer_vto_req / cer_hoy_ref) - 1.0,
        'infla_desconocida': infla_desconocida,
        'f_cer_ult': f_cer_ult, 'cer_ult': cer_ult, 'n_meses': n_meses,
        'tem_be': tem_be, 'tea_be': tea_be, 'tir_real_cer': tir_real_cer,
        'cer_conocido': cer_es_conocido, 'tir_cer_fija': tir_cer_fija,
    }

# =========================================================
# 5. API 1816
# =========================================================
def _get_api_key():
    try:
        k = st.secrets["API_1816_KEY"]
        if k:
            return k
    except Exception:
        pass
    return os.environ.get("API_1816_KEY")

@st.cache_data(ttl=300)
def obtener_precios_1816(lista_tickers, api_key):
    """Lanza excepción si algo falla (así no se cachea el error)."""
    auth = requests.post("https://api.1816.com.ar/v1/auth/token",
                         json={"apiKey": api_key, "module": "mercado"}, timeout=10)
    auth.raise_for_status()
    token = auth.json()["token"]

    url = "https://api.1816.com.ar/v1/mercado/indicadores"
    params = {"tickers": [t.strip().lower() for t in lista_tickers], "campos": ["precioDirty"]}
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=10)
    response.raise_for_status()
    data = response.json()

    precios = {}
    if isinstance(data, dict) and "instrumentos" in data:
        for ticker, valores in data["instrumentos"].items():
            if isinstance(valores, dict) and valores.get("precioDirty") is not None:
                precios[ticker.upper()] = float(valores["precioDirty"])
    if not precios:
        raise ValueError("La API respondió sin precioDirty para ningún ticker.")
    return precios

# =========================================================
# 6. BARRA LATERAL: PRECIOS (EXCEL / API)
# =========================================================
ultima_fila = df_precios.iloc[-1]
fecha_sim = ultima_fila['Fecha_dt']

st.sidebar.header("Parámetros de Simulación")

usar_api = st.sidebar.checkbox("⚡ Usar precios en VIVO (API 1816)", value=False)
precios_vivo = {}
if usar_api:
    api_key = _get_api_key()
    if not api_key:
        st.sidebar.error("Falta API_1816_KEY (en st.secrets o variable de entorno). Usando Excel.")
    else:
        todos = sorted(set([p['ticker_lec'] for p in pares] + [p['ticker_cer'] for p in pares]))
        try:
            with st.sidebar.status("Conectando con 1816..."):
                precios_vivo = obtener_precios_1816(todos, api_key)
            st.sidebar.success(f"✅ Precios 1816 obtenidos ({len(precios_vivo)} especies).")
            fecha_sim = datetime.now(TZ_AR).date()
        except Exception as e:
            st.sidebar.error(f"❌ Falló API: {e}. Usando Excel.")
            precios_vivo = {}

def precio_actual(ticker):
    """Precio API si existe; si no, el de la última fila del Excel. Siempre base 100."""
    v = precios_vivo.get(ticker) if precios_vivo else None
    if v is None or pd.isna(v):
        v = ultima_fila[ticker] if ticker in ultima_fila.index else np.nan
    return normalizar_precio(v)

st.sidebar.write(f"**Fecha del último dato:** {fecha_sim.strftime('%d/%m/%Y')}")
st.sidebar.write(f"**Último CER cargado:** {df_indice['Fecha'].max().strftime('%d/%m/%Y')}")

plazo_opciones = {"T+0 (Contado Inmediato)": 0, "T+1 (24hs)": 1, "T+2 (48hs)": 2}
plazo_label = st.sidebar.selectbox("Plazo de Liquidación:", list(plazo_opciones.keys()), index=1)
plazo_t_selected = plazo_opciones[plazo_label]

st.sidebar.markdown("---")

# Control de unidades: API vs Excel
if precios_vivo:
    with st.expander("🔎 Control de precios: API vs Excel (verificar unidades)"):
        filas = []
        for t in sorted(precios_vivo):
            px_api = normalizar_precio(precios_vivo[t])
            px_xls = normalizar_precio(ultima_fila[t]) if t in ultima_fila.index else np.nan
            filas.append({'Ticker': t, 'API': px_api, 'Excel (última fila)': px_xls,
                          'Dif %': (px_api / px_xls - 1) * 100 if pd.notnull(px_xls) and px_xls else np.nan})
        st.dataframe(pd.DataFrame(filas).style.format(
            {'API': '{:.2f}', 'Excel (última fila)': '{:.2f}', 'Dif %': '{:+.2f}%'}, na_rep="-"),
            use_container_width=True, hide_index=True)

# =========================================================
# 7. TABLA DE RESUMEN
# =========================================================
pares_ordenados = sorted(pares, key=lambda x: x['vencimiento'])

resumen_list = []
omitidos = []
for p in pares_ordenados:
    t_lec, t_cer = p['ticker_lec'], p['ticker_cer']
    px_l, px_c = precio_actual(t_lec), precio_actual(t_cer)
    if pd.isna(px_l) or pd.isna(px_c):
        omitidos.append(f"{t_lec}/{t_cer}")
        continue
    r_all = calcular_be_cer(t_lec, t_cer, px_l, px_c, fecha_sim, df_lec, df_cer, df_indice, feriados,
                            plazo_t=plazo_t_selected, modo_historico=False)
    if r_all:
        resumen_list.append({
            'Vencimiento': p['vencimiento'].strftime('%d/%m/%Y'),
            'Tasa Fija': t_lec,
            'CER': t_cer,
            'TEM BE (%)': r_all['tem_be'] * 100,
            'TEA BE (%)': r_all['tea_be'] * 100,
            'N meses': r_all['n_meses'],
        })

df_resumen = pd.DataFrame(resumen_list)

st.subheader("Resumen de Mercado: Break-Even por Vencimiento")
st.caption(f"Calculado al {fecha_sim.strftime('%d/%m/%Y')} en **{plazo_label}**. "
           "Si el CER del vencimiento ya está publicado no hay inflación que implicar (se indica 'CER conocido').")

if not df_resumen.empty:
    st.dataframe(
        df_resumen.style.format({'TEM BE (%)': '{:.2f}%', 'TEA BE (%)': '{:.2f}%', 'N meses': '{:.2f}'},
                                na_rep="CER conocido"),
        use_container_width=True, hide_index=True
    )
else:
    st.warning("No se pudo calcular ningún par (¿faltan precios o ya vencieron?).")
if omitidos:
    st.caption("Pares omitidos por falta de precio: " + ", ".join(omitidos))

st.markdown("---")

# =========================================================
# 8. DETALLE DE ESPECIE
# =========================================================
opciones_labels = [p['label'] for p in pares_ordenados]
default_idx = next(
    (i for i, p in enumerate(pares_ordenados) if p['ticker_lec'] == 'S30N6' and p['ticker_cer'] == 'X30N6'), 0)

par_seleccionado_label = st.sidebar.selectbox("1. Seleccione el par a comparar:", opciones_labels, index=default_idx)
par_info = next(p for p in pares_ordenados if p['label'] == par_seleccionado_label)
lec_ticker, cer_ticker = par_info['ticker_lec'], par_info['ticker_cer']

def_p_l = precio_actual(lec_ticker)
def_p_c = precio_actual(cer_ticker)
def_p_l = 100.0 if pd.isna(def_p_l) else float(def_p_l)
def_p_c = 100.0 if pd.isna(def_p_c) else float(def_p_c)

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

res_sim = calcular_be_cer(lec_ticker, cer_ticker, precio_lec_sim, precio_cer_sim, fecha_sim,
                          df_lec, df_cer, df_indice, feriados,
                          plazo_t=plazo_t_selected, modo_historico=False)

st.subheader(f"Análisis Detallado: {lec_ticker} vs {cer_ticker} "
             f"(Vto: {par_info['vencimiento'].strftime('%d/%m/%Y')} - {plazo_label})")

if res_sim is None:
    st.warning("No se pudo calcular el par: revisá precios, que no haya vencido, y que el CER cargado "
               "cubra hasta el último dato conocido.")
elif res_sim['cer_conocido']:
    c1, c2, c3 = st.columns(3)
    c1.metric(f"TEA Fija ({lec_ticker})", f"{res_sim['tir_lec']*100:.2f}%")
    c2.metric(f"TIR {cer_ticker} (CER ya conocido)", f"{res_sim['tir_cer_fija']*100:.2f}%")
    c3.metric("TIR CER − TIR Fija", f"{(res_sim['tir_cer_fija'] - res_sim['tir_lec'])*100:+.2f} pp")
    st.info(f"El CER de referencia al vencimiento ({res_sim['f_ref_vto'].strftime('%d/%m/%Y')}) ya está "
            "publicado: la LECER tiene pago conocido y opera como tasa fija. No corresponde calcular una "
            "inflación break-even; se comparan directamente las TIR.")
else:
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(f"TEA Fija ({lec_ticker})", f"{res_sim['tir_lec']*100:.2f}%")
    c2.metric(f"Tasa Real CER ({cer_ticker})", f"{res_sim['tir_real_cer']*100:.2f}%")
    c3.metric(f"Pesos Target ({cer_ticker})", f"${res_sim['pesos_req_cer']:.2f}")
    c4.metric("TEM BE (Mensual)", f"{res_sim['tem_be']*100:.2f}%")
    c5.metric("TEA BE (TEM Anualizada)", f"{res_sim['tea_be']*100:.2f}%")
    st.caption(f"Tramo desconocido: desde {res_sim['f_cer_ult'].strftime('%d/%m/%Y')} hasta "
               f"{res_sim['f_ref_vto'].strftime('%d/%m/%Y')} ({res_sim['n_meses']:.2f} meses efectivos). "
               f"Inflación acumulada implícita en ese tramo: {res_sim['infla_desconocida']*100:.2f}%.")
    if res_sim['n_meses'] < 1.0:
        st.warning("El tramo desconocido es menor a 1 mes: la TEM BE se calcula elevando a 1/N con N chico, "
                   "por lo que pequeñas diferencias de precio la mueven mucho. Mirá la inflación acumulada "
                   "implícita del tramo como referencia más estable.")

# =========================================================
# 9. SERIE HISTÓRICA
# =========================================================
df_precios_filt = df_precios[(df_precios['Fecha_dt'] >= f_desde) & (df_precios['Fecha_dt'] <= f_hasta)]
historico = []

for _, row in df_precios_filt.iterrows():
    f_o = row['Fecha_dt']
    p_l = normalizar_precio(row[lec_ticker]) if lec_ticker in row.index else np.nan
    p_c = normalizar_precio(row[cer_ticker]) if cer_ticker in row.index else np.nan
    if pd.isna(p_l) or pd.isna(p_c):
        continue

    r = calcular_be_cer(lec_ticker, cer_ticker, p_l, p_c, f_o, df_lec, df_cer, df_indice, feriados,
                        plazo_t=plazo_t_selected, modo_historico=True)
    if r:
        if not df_tamar.empty:
            tamar_sub = df_tamar[df_tamar['Fecha_dt'] <= f_o]
            r['tamar'] = tamar_sub['Tamar_pct'].iloc[-1] if not tamar_sub.empty else np.nan
        else:
            r['tamar'] = np.nan
        historico.append(r)

df_hist = pd.DataFrame(historico)

# =========================================================
# 10. GRÁFICOS
# =========================================================
st.markdown("---")
st.subheader("Evolución Histórica de Métricas")

if not df_hist.empty and df_hist['tem_be'].notna().any():
    df_hist['tem_be_pct'] = df_hist['tem_be'] * 100
    df_hist['tea_be_pct'] = df_hist['tea_be'] * 100
    df_hist['tir_real_cer_pct'] = df_hist['tir_real_cer'] * 100
    st.caption("Los días en que el CER del vencimiento ya era conocido no se grafican (no hay BE).")

    fig1 = go.Figure()
    fig1.add_trace(go.Scatter(
        x=df_hist['fecha_op'], y=df_hist['tem_be_pct'],
        name="TEM BE (%)", line=dict(color='#d62728', width=2.5),
        hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEM BE:</b> %{y:.2f}%<extra></extra>"
    ))
    fig1.update_layout(
        title=f"<b>1. Inflación Mensual Implícita (TEM BE): {lec_ticker} vs {cer_ticker}</b>",
        xaxis_title="Fecha", yaxis_title="TEM (%)", template="plotly_white", hovermode="x unified")
    st.plotly_chart(fig1, use_container_width=True)

    fig2 = make_subplots(specs=[[{"secondary_y": True}]])
    fig2.add_trace(go.Scatter(
        x=df_hist['fecha_op'], y=df_hist['tea_be_pct'], name="TEA BE (%)",
        line=dict(color='#1f77b4', width=2.5),
        hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEA BE:</b> %{y:.2f}%<extra></extra>"),
        secondary_y=False)
    fig2.add_trace(go.Scatter(
        x=df_hist['fecha_op'], y=df_hist['tir_real_cer_pct'], name=f"Tasa Real CER ({cer_ticker})",
        line=dict(color='#2ca02c', width=2.5),
        hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Tasa Real CER:</b> %{y:.2f}%<extra></extra>"),
        secondary_y=True)
    fig2.update_layout(
        title=f"<b>2. Inflación Anualizada Implícita (TEA BE) vs Tasa Real CER ({cer_ticker})</b>",
        template="plotly_white", hovermode="x unified")
    fig2.update_xaxes(title_text="Fecha")
    fig2.update_yaxes(title_text="<b>TEA BE (%)</b>", secondary_y=False, title_font_color='#1f77b4')
    fig2.update_yaxes(title_text="<b>Tasa Real CER (%)</b>", secondary_y=True, title_font_color='#2ca02c')
    st.plotly_chart(fig2, use_container_width=True)

    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
    fig3.add_trace(go.Scatter(
        x=df_hist['fecha_op'], y=df_hist['tea_be_pct'], name="TEA BE (%)",
        line=dict(color='#1f77b4', width=2.5),
        hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEA BE:</b> %{y:.2f}%<extra></extra>"),
        secondary_y=False)
    fig3.add_trace(go.Scatter(
        x=df_hist['fecha_op'], y=df_hist['tamar'], name="Tasa TAMAR (%)",
        line=dict(color='#ff7f0e', width=2.5),
        hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR:</b> %{y:.2f}%<extra></extra>"),
        secondary_y=True)
    fig3.update_layout(title="<b>3. Inflación Anualizada Implícita (TEA BE) vs Tasa TAMAR</b>",
                       template="plotly_white", hovermode="x unified")
    fig3.update_xaxes(title_text="Fecha")
    fig3.update_yaxes(title_text="<b>TEA BE (%)</b>", secondary_y=False, title_font_color='#1f77b4')
    fig3.update_yaxes(title_text="<b>Tasa TAMAR (%)</b>", secondary_y=True, title_font_color='#ff7f0e')
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.warning("No hay datos con inflación break-even para el rango seleccionado "
               "(o el CER de todo el período ya era conocido al vencimiento).")