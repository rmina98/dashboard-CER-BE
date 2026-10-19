import os
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
import requests

# Configuración de la página
st.set_page_config(page_title="Dashboard TAMAR BE", layout="wide")
st.title("📊 Dashboard: TAMAR Break-Even")

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# Overrides manuales de VPV (solo si no está en la columna 'VPV' del Excel).
# Lo ideal es pasar esto a la columna VPV de la hoja Lecaps y dejar el dict vacío.
VPV_OVERRIDE = {'S29E7': 111.6660}

# =========================================================
# 📂 1. CARGA Y PREPARACIÓN DE DATOS
# =========================================================
@st.cache_data
def cargar_excel():
    if not os.path.exists("Inputs.xlsx"):
        st.error("❌ Archivo 'Inputs.xlsx' no encontrado en la carpeta.")
        st.stop()

    xls = pd.ExcelFile("Inputs.xlsx")
    df_lec = pd.read_excel(xls, sheet_name='Lecaps')
    df_let = pd.read_excel(xls, sheet_name='Letam')
    df_precios = pd.read_excel(xls, sheet_name='Precios')
    df_tamar = pd.read_excel(xls, sheet_name='Tamar')
    df_feriados = pd.read_excel(xls, sheet_name='Feriados')

    feriados = set(pd.to_datetime(df_feriados["Feriados"]).dt.date)

    df_tamar['Fecha'] = pd.to_datetime(df_tamar['Fecha']).dt.date
    df_tamar = df_tamar[df_tamar['Tamar'] > 0].sort_values('Fecha').reset_index(drop=True)

    df_precios['Fecha_dt'] = pd.to_datetime(df_precios['Fecha']).dt.date
    df_precios = df_precios.sort_values('Fecha_dt').reset_index(drop=True)

    return df_lec, df_let, df_precios, df_tamar, feriados

df_lec, df_let, df_precios, df_tamar, feriados = cargar_excel()

# =========================================================
# 📅 2. FUNCIONES DE CALENDARIO
# =========================================================
def es_habil(f, feriados):
    return f.weekday() < 5 and f not in feriados

def proximo_habil(f, feriados):
    actual = f
    while not es_habil(actual, feriados):
        actual += timedelta(days=1)
    return actual

def restar_dias_habiles(f, dias, feriados):
    actual = f
    restados = 0
    while restados < dias:
        actual -= timedelta(days=1)
        if es_habil(actual, feriados):
            restados += 1
    return actual

def contar_dias_habiles(f_inicio, f_fin, feriados):
    """Cuenta días hábiles en el intervalo (f_inicio, f_fin]: excluye inicio, incluye fin."""
    actual = f_inicio
    dias = 0
    while actual < f_fin:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            dias += 1
    return dias

def calcular_fecha_liq(f_op, plazo_t, feriados):
    actual = f_op
    sumados = 0
    while sumados < plazo_t:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            sumados += 1
    return proximo_habil(actual, feriados)

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

# =========================================================
# 🧮 3. MOTOR DE CÁLCULO
# =========================================================
def normalizar_precio(x):
    """Precios pueden venir como 1.05 o 105. Devuelve siempre en base 100. NaN-safe."""
    if x is None or pd.isna(x):
        return np.nan
    x = float(x)
    return x * 100 if x < 10 else x

def tem_desde_tasa_comb(tasa_comb):
    """TAMAR TEM según prospecto, con tasa_comb = TAMAR + MARGEN (TNA decimal)."""
    return ((1 + tasa_comb / (365 / 32)) ** (365 / 32)) ** (1 / 12) - 1

def calcular_tamar_be(ticker_lec, ticker_let, p_lec, p_let, f_op,
                      df_lec, df_let, df_tamar, feriados, plazo_t=1):
    """Devuelve dict de resultados, o None si el cálculo no es válido para esa fecha."""
    if pd.isna(p_lec) or pd.isna(p_let) or p_lec <= 0 or p_let <= 0:
        return None

    row_lec = df_lec[df_lec['Ticker'] == ticker_lec].iloc[0]
    row_let = df_let[df_let['Ticker'] == ticker_let].iloc[0]
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
        return None  # Lecap ya vencida / liquidando después del cobro

    rend_lec = (vpv_lec / p_lec) - 1
    tir_lec = ((1 + rend_lec) ** (365 / dias_cartera_lec)) - 1
    tna_lec = rend_lec * (365 / dias_cartera_lec)

    # --- LETAM ---
    f_vto_let = pd.to_datetime(row_let['Vencimiento']).date()
    f_em_let = pd.to_datetime(row_let['Emisión']).date()
    margen_let = row_let['Tasa']
    d360_let = dias360_excel(f_em_let, f_vto_let)
    t_meses_let = (d360_let / 360) * 12
    f_cobro_let = proximo_habil(f_vto_let, feriados)

    dias_cartera_let = (f_cobro_let - f_liq).days
    if dias_cartera_let <= 0:
        return None

    # VPV que debería tener la Letam para igualar la TIR de la Lecap
    # (supone reinversión a la TIR de la Lecap si ésta vence antes)
    vpv_objetivo = p_let * ((1 + tir_lec) ** (dias_cartera_let / 365))

    tem_cupon_target = ((vpv_objetivo / 100.0) ** (1.0 / t_meses_let)) - 1
    tea_target = (1 + tem_cupon_target) ** 12
    tasa_comb_target = ((tea_target ** (32 / 365)) - 1) * (365 / 32)  # = TAMAR_prom + MARGEN

    # --- FIXING TAMAR: ventana [emisión-10h, vto-10h], ambos inclusive ---
    f_inicio_fixing = restar_dias_habiles(f_em_let, 10, feriados)
    f_fin_fixing = restar_dias_habiles(f_vto_let, 10, feriados)
    n_total = contar_dias_habiles(f_inicio_fixing - timedelta(days=1), f_fin_fixing, feriados)

    f_corte_dev = min(f_op, f_fin_fixing)
    if f_corte_dev >= f_inicio_fixing:
        n_dev = contar_dias_habiles(f_inicio_fixing - timedelta(days=1), f_corte_dev, feriados)
        df_dev = df_tamar[(df_tamar['Fecha'] >= f_inicio_fixing) & (df_tamar['Fecha'] <= f_corte_dev)]
        dev_tna = df_dev['Tamar'].mean() / 100.0 if not df_dev.empty else np.nan
    else:
        n_dev, dev_tna = 0, 0.0

    n_proy = n_total - n_dev

    if n_proy > 0:
        # promedio = (n_dev*dev + n_proy*BE) / n_total ; tasa_comb = margen + promedio
        tamar_be = ((tasa_comb_target - margen_let) * n_total - n_dev * dev_tna) / n_proy
    else:
        tamar_be = np.nan

    if n_proy > 0 and np.isnan(tamar_be):
        return None  # faltan datos de TAMAR dentro de la ventana

    # TAMAR de mercado vigente a f_op (última publicada)
    df_hist_mercado = df_tamar[df_tamar['Fecha'] <= f_op]
    if df_hist_mercado.empty:
        return None
    tamar_mercado = df_hist_mercado['Tamar'].iloc[-1] / 100.0

    spread = (tamar_mercado - tamar_be) * 100 if n_proy > 0 else 0.0

    # Si el fixing ya cerró, la Letam es efectivamente de tasa fija: calculamos su TIR
    tir_let = np.nan
    if n_proy == 0 and not np.isnan(dev_tna):
        tem_fija = tem_desde_tasa_comb(margen_let + dev_tna)
        vpv_let = 100.0 * (1 + tem_fija) ** t_meses_let
        rend_let = vpv_let / p_let - 1
        tir_let = (1 + rend_let) ** (365 / dias_cartera_let) - 1

    return {
        'fecha': f_op, 'tir_lec': tir_lec, 'tna_lec': tna_lec, 'vpv_objetivo': vpv_objetivo,
        'tamar_be': tamar_be, 'tamar_mercado': tamar_mercado, 'spread_pkt': spread,
        'es_fija': n_proy == 0, 'tir_let': tir_let,
        # auxiliares (para el test)
        'dias_cartera_let': dias_cartera_let, 't_meses_let': t_meses_let, 'margen_let': margen_let,
        'f_inicio_fixing': f_inicio_fixing, 'f_fin_fixing': f_fin_fixing,
        'n_total': n_total, 'n_dev': n_dev, 'n_proy': n_proy,
    }

# =========================================================
# 🧪 4. TEST SINTÉTICO DEL MOTOR
# =========================================================
def test_sintetico(ticker_lec, ticker_let, p_lec, x_tamar, df_lec, df_let, feriados):
    """
    TAMAR constante = x_tamar (decimal) en toda la ventana. Se calcula el precio de la Letam
    que da exactamente la TIR de la Lecap bajo esa TAMAR. El BE debe devolver x_tamar.
    """
    row_let = df_let[df_let['Ticker'] == ticker_let].iloc[0]
    f_em = pd.to_datetime(row_let['Emisión']).date()
    f_vto = pd.to_datetime(row_let['Vencimiento']).date()
    ini = restar_dias_habiles(f_em, 10, feriados)
    fin = restar_dias_habiles(f_vto, 10, feriados)

    # TAMAR sintética constante, por día hábil, con margen de sobra
    dias = []
    d = ini - timedelta(days=30)
    while d <= fin + timedelta(days=30):
        if es_habil(d, feriados):
            dias.append(d)
        d += timedelta(days=1)
    df_syn = pd.DataFrame({'Fecha': dias, 'Tamar': x_tamar * 100.0})

    f_op = proximo_habil(ini + (fin - ini) / 2, feriados)

    # 1ra pasada: obtener TIR Lecap y datos de la Letam
    r0 = calcular_tamar_be(ticker_lec, ticker_let, p_lec, 100.0, f_op, df_lec, df_let, df_syn, feriados)
    if r0 is None:
        return None, "No se pudo calcular el caso base (¿fechas/datos fuera de rango?)."

    tem = tem_desde_tasa_comb(r0['margen_let'] + x_tamar)
    vpv_x = 100.0 * (1 + tem) ** r0['t_meses_let']
    p_eq = vpv_x / ((1 + r0['tir_lec']) ** (r0['dias_cartera_let'] / 365))

    # 2da pasada: con el precio de equilibrio, el BE debe ser x_tamar
    r1 = calcular_tamar_be(ticker_lec, ticker_let, p_lec, p_eq, f_op, df_lec, df_let, df_syn, feriados)
    if r1 is None:
        return None, "No se pudo calcular el caso de equilibrio."
    return r1, f_op

# =========================================================
# 🔌 5. API 1816 (PRECIOS EN VIVO)
# =========================================================
def _k(ticker):
    """Clave normalizada de ticker (la API devuelve en mayúsculas)."""
    return str(ticker).strip().upper()

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
# 🎛️ 6. INTERFAZ DE USUARIO (DASHBOARD)
# =========================================================
st.sidebar.header("⚙️ Parámetros de Simulación")

usar_api = st.sidebar.checkbox("⚡ Usar precios en VIVO (API 1816)", value=False)
precios_vivo = {}
if usar_api:
    api_key = _get_api_key()
    if not api_key:
        st.sidebar.error("Falta API_1816_KEY (en st.secrets o variable de entorno). Usando Excel.")
    else:
        todos = sorted({_k(t) for t in df_lec['Ticker']} | {_k(t) for t in df_let['Ticker']})
        try:
            with st.sidebar.status("Conectando con 1816..."):
                precios_vivo = obtener_precios_1816(todos, api_key)
            st.sidebar.success(f"✅ Precios 1816 obtenidos ({len(precios_vivo)} especies).")
        except Exception as e:
            st.sidebar.error(f"❌ Falló API: {e}. Usando Excel.")
            precios_vivo = {}

lecap_elegida = st.sidebar.selectbox("1️⃣ Tasa Fija (Lecap/Boncap):", df_lec['Ticker'].tolist())
letam_elegida = st.sidebar.selectbox("2️⃣ Tasa Variable (Letam):", df_let['Ticker'].tolist())

px_l_api = precios_vivo.get(_k(lecap_elegida))
px_t_api = precios_vivo.get(_k(letam_elegida))
usa_vivo = px_l_api is not None and px_t_api is not None

if usa_vivo:
    def_p_l = float(normalizar_precio(px_l_api))
    def_p_t = float(normalizar_precio(px_t_api))
    fecha_sim = datetime.now(TZ_AR).date()
    st.sidebar.write(f"**Precios en vivo al:** {fecha_sim.strftime('%d/%m/%Y')}")
else:
    if usar_api and precios_vivo:
        st.sidebar.warning("La API no devolvió precio para este par. Usando Excel.")
    # Última fecha con precio válido para ambas especies
    df_validos = df_precios.dropna(subset=[lecap_elegida, letam_elegida])
    if df_validos.empty:
        st.error("❌ No hay fechas con precio para ambas especies seleccionadas.")
        st.stop()
    ultima_fila = df_validos.iloc[-1]
    fecha_sim = ultima_fila['Fecha_dt']
    def_p_l = float(normalizar_precio(ultima_fila[lecap_elegida]))
    def_p_t = float(normalizar_precio(ultima_fila[letam_elegida]))
    st.sidebar.write(f"**Última fecha con ambos precios:** {fecha_sim.strftime('%d/%m/%Y')}")

st.sidebar.markdown("---")
precio_lecap_sim = st.sidebar.number_input("Precio Lecap:", value=def_p_l, step=0.10, format="%.2f")
precio_letam_sim = st.sidebar.number_input("Precio Letam:", value=def_p_t, step=0.10, format="%.2f")

st.sidebar.markdown("---")

# Selector de Rango de Fechas para la Histórica
st.sidebar.subheader("📅 Rango Histórico a Graficar")
fecha_min_general = df_precios['Fecha_dt'].min()
fecha_max_general = df_precios['Fecha_dt'].max()

rango_fechas = st.sidebar.date_input(
    "Seleccioná el período:",
    value=(fecha_min_general, fecha_max_general),
    min_value=fecha_min_general,
    max_value=fecha_max_general
)

if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
    f_desde, f_hasta = rango_fechas
else:
    f_desde, f_hasta = fecha_min_general, fecha_max_general

# Control de unidades: API vs Excel
if precios_vivo:
    with st.expander("🔎 Control de precios: API vs Excel (verificar unidades)"):
        ult_xls = df_precios.iloc[-1]
        filas = []
        for t in sorted(precios_vivo):
            px_api = normalizar_precio(precios_vivo[t])
            col = next((c for c in ult_xls.index if _k(c) == t), None)
            px_xls = normalizar_precio(ult_xls[col]) if col is not None else np.nan
            filas.append({'Ticker': t, 'API': px_api, 'Excel (última fila)': px_xls,
                          'Dif %': (px_api / px_xls - 1) * 100 if pd.notnull(px_xls) and px_xls else np.nan})
        st.dataframe(pd.DataFrame(filas).style.format(
            {'API': '{:.2f}', 'Excel (última fila)': '{:.2f}', 'Dif %': '{:+.2f}%'}, na_rep="-"),
            use_container_width=True, hide_index=True)

# =========================================================
# 🚀 7. EJECUCIÓN Y RENDERIZADO
# =========================================================
res_sim = calcular_tamar_be(lecap_elegida, letam_elegida, precio_lecap_sim, precio_letam_sim, fecha_sim,
                            df_lec, df_let, df_tamar, feriados)

if res_sim is None:
    st.error("⚠️ No se pudo calcular la simulación: revisá que las especies no hayan vencido a la fecha "
             "y que haya datos de TAMAR para la ventana de fixing.")
    st.stop()

st.caption(f"Simulación al {fecha_sim.strftime('%d/%m/%Y')} con precios "
           f"{'en VIVO (1816)' if usa_vivo else 'del Excel'}.")

col1, col2, col3 = st.columns(3)
col1.metric(label=f"TNA {lecap_elegida} (Fija)", value=f"{res_sim['tna_lec']*100:.2f}%")
col2.metric(label=f"VPV Req {letam_elegida} (Var)", value=f"${res_sim['vpv_objetivo']:.2f}")

if res_sim['es_fija']:
    col3.metric(label="Spread TAMAR BE vs Mkt", value="Ya es Fija")
    st.markdown("---")
    st.info("⚠️ **El período de fixing finalizó:** la Letam ya devengó toda la TAMAR y opera como "
            "instrumento a tasa fija, por lo que no corresponde calcular una Break-Even futura.")
    if not np.isnan(res_sim['tir_let']):
        ca, cb = st.columns(2)
        ca.info(f"**TIR {letam_elegida} (ya conocida):** {res_sim['tir_let']*100:.2f}%")
        cb.info(f"**TIR {lecap_elegida}:** {res_sim['tir_lec']*100:.2f}%")
else:
    col3.metric(label="Spread TAMAR BE vs Mkt", value=f"{res_sim['spread_pkt']:+.2f} pkt")
    st.markdown("---")
    col_a, col_b = st.columns(2)
    col_a.info(f"**TAMAR Break-Even Futura:** {res_sim['tamar_be']*100:.2f}%")
    col_b.info(f"**TAMAR Mercado Hoy:** {res_sim['tamar_mercado']*100:.2f}%")
    st.caption(
        f"Ventana de fixing: {res_sim['f_inicio_fixing']:%d/%m/%Y} a {res_sim['f_fin_fixing']:%d/%m/%Y} "
        f"({res_sim['n_total']} hábiles; {res_sim['n_dev']} conocidos, {res_sim['n_proy']} a proyectar). "
        "El BE es la TAMAR constante sobre los días a proyectar que iguala la TIR de la Lecap. "
        "Supone reinversión a la TIR de la Lecap si ésta vence antes que la Letam."
    )

# ---------------------------------------------------------
# Histórica (siempre con precios del Excel)
# ---------------------------------------------------------
df_precios_filtrado = df_precios[(df_precios['Fecha_dt'] >= f_desde) & (df_precios['Fecha_dt'] <= f_hasta)]

historico = []
for _, row in df_precios_filtrado.iterrows():
    p_l = normalizar_precio(row[lecap_elegida])
    p_t = normalizar_precio(row[letam_elegida])
    if pd.isna(p_l) or pd.isna(p_t):
        continue
    r = calcular_tamar_be(lecap_elegida, letam_elegida, p_l, p_t, row['Fecha_dt'],
                          df_lec, df_let, df_tamar, feriados)
    if r is not None:
        historico.append(r)

df_hist = pd.DataFrame(historico)

st.subheader("📈 Evolución Histórica: BE vs Mercado")

if not df_hist.empty:
    # Se excluyen los días con fixing cerrado para que la línea no caiga a 0% arruinando el eje Y
    df_hist_plot = df_hist[~df_hist['es_fija']].copy()

    if not df_hist_plot.empty:
        df_hist_plot['be_pct'] = df_hist_plot['tamar_be'] * 100
        df_hist_plot['mkt_pct'] = df_hist_plot['tamar_mercado'] * 100

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_hist_plot['fecha'], y=df_hist_plot['be_pct'], mode='lines',
            name=f'TAMAR BE ({lecap_elegida} vs {letam_elegida})',
            line=dict(color='#d62728', width=2.5),
            hovertemplate='<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR BE:</b> %{y:.2f}%<extra></extra>'
        ))
        fig.add_trace(go.Scatter(
            x=df_hist_plot['fecha'], y=df_hist_plot['mkt_pct'], mode='lines',
            name='TAMAR Mkt Real',
            line=dict(color='#1f77b4', width=2.5, dash='dash'),
            hovertemplate='<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR Mkt Real:</b> %{y:.2f}%<extra></extra>'
        ))
        fig.update_layout(
            hovermode='x unified', yaxis_title='TNA (%)', template='plotly_white',
            height=450, margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("⚠️ Todos los datos en este rango corresponden al período donde la Letam ya operaba "
                   "como tasa fija (cierre de fixing).")
else:
    st.warning("No hay datos válidos en el rango de fechas seleccionado.")

# ---------------------------------------------------------
# Test de validación del motor
# ---------------------------------------------------------
with st.expander("🧪 Test de validación del motor (TAMAR sintética constante)"):
    x_test = st.number_input("TAMAR constante a testear (TNA %):", value=30.0, step=1.0, format="%.2f") / 100.0
    if st.button("Correr test"):
        r_test, info = test_sintetico(lecap_elegida, letam_elegida, precio_lecap_sim, x_test,
                                      df_lec, df_let, feriados)
        if r_test is None:
            st.error(info)
        else:
            err = abs(r_test['tamar_be'] - x_test) * 100
            st.write(f"Fecha de operación del test: {info:%d/%m/%Y} "
                     f"(n_total={r_test['n_total']}, n_dev={r_test['n_dev']}, n_proy={r_test['n_proy']})")
            st.write(f"BE obtenido: **{r_test['tamar_be']*100:.6f}%** vs esperado **{x_test*100:.6f}%** "
                     f"(error: {err:.2e} pp)")
            if err < 1e-6:
                st.success("✅ El motor reproduce la TAMAR constante.")
            else:
                st.error("❌ El BE no coincide con la TAMAR sintética: hay un problema en el motor.")