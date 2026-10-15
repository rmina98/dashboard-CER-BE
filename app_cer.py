# =========================================================
# 3. TABLA DE RESUMEN GENERAL E INTEGRACIÓN CON 1816 API
# =========================================================
# Función protegida por caché (5 min) para no gastar de más los créditos
@st.cache_data(ttl=300)
def obtener_precios_1816(lista_tickers):
    api_key_1816 = "25a792a5ce74054a5bed8a559c27f98aa651472a59d73810e33553feee0b31ef"
    
    try:
        # 1. Obtener Token
        auth = requests.post(
            "https://api.1816.com.ar/v1/auth/token",
            json={"apiKey": api_key_1816, "module": "mercado"},
            timeout=10
        )
        auth.raise_for_status()
        token = auth.json()["token"]

        # 2. Consultar Precios
        url = "https://api.1816.com.ar/v1/mercado/indicadores"
        tickers_limpios = [t.strip().lower() for t in lista_tickers]
        
        params = {
            "tickers": tickers_limpios,
            "campos": ["precioDirty"]
        }
        
        response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        
        # 3. Formatear la respuesta (Ajustado según el recuadro amarillo)
        precios_dict = {}
        
        # 1816 envía un diccionario principal que contiene 'instrumentos'
        if isinstance(data, dict) and "instrumentos" in data:
            instrumentos = data["instrumentos"]
            for ticker, valores in instrumentos.items():
                if isinstance(valores, dict) and "precioDirty" in valores:
                    precios_dict[ticker.upper()] = float(valores["precioDirty"])
                    
        # Si la API respondió 200 OK pero no encontró datos
        if not precios_dict:
            st.sidebar.error(f"⚠️ No se pudieron extraer los precios. Data cruda: {data}")
            return None
            
        return precios_dict

    except requests.exceptions.HTTPError as errh:
        st.sidebar.error(f"🛑 Error HTTP: {errh.response.status_code} - {errh.response.text}")
        return None
    except Exception as e:
        st.sidebar.error(f"⚠️ Error interno en la consulta: {str(e)}")
        return None


# Clonamos la última fila para que sea editable
ultima_fila = df_precios.iloc[-1].copy()
fecha_sim = ultima_fila['Fecha_dt']

st.sidebar.header("Parámetros de Simulación")

# ---- CHECKBOX MAGICO DE 1816 ----
usar_api = st.sidebar.checkbox("⚡ Usar precios en VIVO (API 1816)", value=False)

if usar_api:
    todos_los_tickers = list(set([p['ticker_lec'] for p in pares] + [p['ticker_cer'] for p in pares]))
    
    with st.sidebar.status("Conectando con 1816...", expanded=False) as status:
        precios_vivo = obtener_precios_1816(todos_los_tickers)
        if precios_vivo:
            status.update(label="✅ Precios obtenidos de 1816", state="complete", expanded=False)
            fecha_sim = datetime.now().date() 
            for t_api, px_api in precios_vivo.items():
                ultima_fila[t_api] = px_api
        else:
            status.update(label="❌ Falló la conexión API, usando Excel", state="error", expanded=True)
# ---------------------------------

st.sidebar.write(f"**Fecha del precio utilizado:** {fecha_sim.strftime('%d/%m/%Y')}")

# Selector de Plazo de Liquidación
plazo_opciones = {"T+0 (Contado Inmediato)": 0, "T+1 (24hs)": 1, "T+2 (48hs)": 2}
plazo_label = st.sidebar.selectbox("Plazo de Liquidación:", list(plazo_opciones.keys()), index=1)
plazo_t_selected = plazo_opciones[plazo_label]

st.sidebar.markdown("---")

# ¡AQUÍ ASEGURAMOS QUE pares_ordenados EXISTA SIEMPRE!
pares_ordenados = sorted(pares, key=lambda x: x['vencimiento'])

resumen_list = []
for p in pares_ordenados:
    t_lec, t_cer = p['ticker_lec'], p['ticker_cer']
    val_l_r = ultima_fila[t_lec] if t_lec in ultima_fila else None
    val_c_r = ultima_fila[t_cer] if t_cer in ultima_fila else None
    
    if pd.notnull(val_l_r) and pd.notnull(val_c_r):
        px_l = val_l_r * 100 if val_l_r < 10 else val_l_r
        px_c = val_c_r * 100 if val_c_r < 10 else val_c_r
        r_all = calcular_be_cer(t_lec, t_cer, px_l, px_c, fecha_sim, df_lec, df_cer, df_indice, feriados, plazo_t=plazo_t_selected)
        if r_all:
            resumen_list.append({
                'Vencimiento_dt': p['vencimiento'],
                'Vencimiento': p['vencimiento'].strftime('%d/%m/%Y'),
                'Tasa Fija': t_lec,
                'CER': t_cer,
                'TEM BE (%)': r_all['tem_be'] * 100,
                'TEA BE (%)': r_all['tea_be'] * 100
            })

df_resumen = pd.DataFrame(resumen_list)

st.subheader("Resumen de Mercado: Break-Even por Vencimiento")
origen_texto = "API 1816 (Tiempo Real)" if usar_api else "Excel"
st.caption(f"Calculado con origen de datos: **{origen_texto}** al {fecha_sim.strftime('%d/%m/%Y')} en **{plazo_label}**:")

if not df_resumen.empty:
    df_tabla_show = df_resumen[['Vencimiento', 'Tasa Fija', 'CER', 'TEM BE (%)', 'TEA BE (%)']]
    st.dataframe(
        df_tabla_show.style.format({
            'TEM BE (%)': '{:.2f}%',
            'TEA BE (%)': '{:.2f}%'
        }),
        use_container_width=True,
        hide_index=True
    )

st.markdown("---")