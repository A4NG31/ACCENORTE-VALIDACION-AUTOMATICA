import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD =====
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

# Monkey patch para evitar problemas de watcher
import streamlit.web.bootstrap
import streamlit.watcher

def no_op_watch(*args, **kwargs):
    return lambda: None

def no_op_watch_file(*args, **kwargs):
    return

streamlit.watcher.path_watcher.watch_file = no_op_watch_file
streamlit.watcher.path_watcher._watch_path = no_op_watch
streamlit.watcher.event_based_path_watcher.EventBasedPathWatcher.__init__ = lambda *args, **kwargs: None
streamlit.web.bootstrap._install_config_watchers = lambda *args, **kwargs: None

# ===== IMPORTS NORMALES =====
import streamlit as st
import pandas as pd
import re
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Configuración Streamlit
st.set_page_config(
    page_title="Validador Power BI - ACCENORTE",
    page_icon="💰",
    layout="wide"
)

# ===== CSS PERSONALIZADO =====
st.markdown("""
<style>
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1E1E2F 0%, #2D2D44 100%);
        color: white;
    }
    
    [data-testid="stSidebar"] .stMarkdown, 
    [data-testid="stSidebar"] .stText, 
    [data-testid="stSidebar"] .stInfo {
        color: white !important;
    }
    
    /* Main content styling */
    .main-header {
        background: linear-gradient(90deg, #00C9FF 0%, #92FE9D 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        font-size: 2.5em;
        font-weight: bold;
        margin-bottom: 1em;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        padding: 20px;
        color: white;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .success-box {
        background: linear-gradient(135deg, #00b09b 0%, #96c93d 100%);
        color: white;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        text-align: center;
        font-weight: bold;
        font-size: 1.2em;
    }
    
    .error-box {
        background: linear-gradient(135deg, #ff416c 0%, #ff4b2b 100%);
        color: white;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        text-align: center;
        font-weight: bold;
        font-size: 1.2em;
    }
    
    .warning-box {
        background: linear-gradient(135deg, #f7971e 0%, #ffd200 100%);
        color: white;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        text-align: center;
        font-weight: bold;
        font-size: 1.2em;
    }
    
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #00C9FF 0%, #92FE9D 100%);
    }
    
    .stSpinner > div > div {
        border-color: #00C9FF;
    }
</style>
""", unsafe_allow_html=True)

# ===== LOGO Y HEADER =====
st.markdown("""
<div style="text-align: center; margin-bottom: 40px;">
    <h1 class="main-header">💰 Validador Power BI - ACCENORTE</h1>
    <p style="color: #666; font-size: 1.2em;">Sistema de validación y conciliación automática</p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ===== FUNCIONES PRINCIPALES =====

def extraer_fecha_desde_excel(uploaded_file):
    """Extrae la fecha desde la celda combinada (G18:N24) del Excel"""
    try:
        df = pd.read_excel(uploaded_file, header=None)
        
        for fila in range(17, 24):
            for col in range(6, 14):
                if pd.notna(df.iloc[fila, col]):
                    celda = str(df.iloc[fila, col]).strip()
                    patrones_fecha = [
                        r'(\d{1,2})/(\d{1,2})/(\d{4})',
                        r'(\d{1,2})-(\d{1,2})-(\d{4})',
                        r'(\d{4})-(\d{1,2})-(\d{1,2})'
                    ]
                    
                    for patron in patrones_fecha:
                        match = re.search(patron, celda)
                        if match:
                            if '/' in celda:
                                dia, mes, año = match.groups()
                            elif '-' in celda and len(match.group(1)) == 4:
                                año, mes, dia = match.groups()
                            else:
                                dia, mes, año = match.groups()
                            
                            fecha = datetime(int(año), int(mes), int(dia))
                            st.success(f"📅 Fecha detectada: {fecha.strftime('%d/%m/%Y')}")
                            return fecha.strftime("%Y-%m-%d")
        
        st.error("❌ No se encontró fecha en el rango G18:N24")
        return None
        
    except Exception as e:
        st.error(f"❌ Error al extraer fecha del Excel: {e}")
        return None

def procesar_excel(uploaded_file):
    """Procesa el archivo Excel para extraer valor a pagar y número de pasos"""
    try:
        df = pd.read_excel(uploaded_file, header=None)
        
        valor_a_pagar = 0
        numero_pasos = 0
        
        # Buscar fila con "Valor" en columna AK (índice 36)
        for idx, fila in df.iterrows():
            if pd.notna(fila[36]) and str(fila[36]).strip().upper() == "VALOR":
                for i in range(idx + 1, len(df)):
                    valor_celda = df.iloc[i, 36]
                    if pd.notna(valor_celda):
                        try:
                            valor_num = float(valor_celda)
                            valor_a_pagar += valor_num
                        except:
                            continue
                break
        
        # Buscar "TOTAL TRANSACCIONES"
        for idx, fila in df.iterrows():
            for col in range(len(fila)):
                celda = str(fila[col])
                if "TOTAL TRANSACCIONES" in celda.upper():
                    numeros = re.findall(r'\d+', celda)
                    if numeros:
                        numero_pasos = int(numeros[0])
                        break
            if numero_pasos > 0:
                break
        
        return valor_a_pagar, numero_pasos
        
    except Exception as e:
        st.error(f"❌ Error procesando Excel: {e}")
        return 0, 0

def setup_driver():
    """Configurar ChromeDriver"""
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        st.error(f"❌ Error configurando ChromeDriver: {e}")
        return None

def click_conciliacion_date(driver, fecha_objetivo):
    """Hacer clic en la conciliación específica por fecha"""
    try:
        fecha_formateada = f"{fecha_objetivo} 00:00 al {fecha_objetivo} 11:59"
        
        st.info(f"🔍 Buscando conciliación: {fecha_formateada}")
        
        # Esperar a que carguen los elementos
        time.sleep(8)
        
        selectors = [
            f"//*[contains(text(), 'Conciliación Accenorte del {fecha_formateada}')]",
            f"//*[contains(text(), 'CONCILIACIÓN ACCENORTE DEL {fecha_formateada}')]",
            f"//*[contains(text(), '{fecha_formateada}')]",
            f"//*[contains(text(), 'Conciliación Accenorte')]",
        ]
        
        elemento_conciliacion = None
        for selector in selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        elemento_conciliacion = elemento
                        st.success("✅ Conciliación encontrada")
                        break
                if elemento_conciliacion:
                    break
            except:
                continue
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_conciliacion)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(5)
            return True
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            return False
            
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliación: {str(e)}")
        return False

def find_accenorte_data(driver):
    """Extrae los valores de VALOR A PAGAR A COMERCIO y CANTIDAD PASOS"""
    try:
        valor_a_pagar = None
        cantidad_pasos = None
        
        st.info("🔍 Extrayendo datos del Power BI...")
        time.sleep(5)
        
        # Buscar en elementos de la esquina superior izquierda
        elementos_esquina = driver.find_elements(By.XPATH, "//*[position() < 30]")
        
        for elemento in elementos_esquina:
            if elemento.is_displayed():
                location = elemento.location
                if location['x'] < 800 and location['y'] < 800:
                    texto_completo = elemento.text.strip()
                    if texto_completo and len(texto_completo) > 10:
                        
                        # Reconstruir texto eliminando espacios entre letras
                        texto_reconstruido = ""
                        palabras = texto_completo.split()
                        
                        i = 0
                        while i < len(palabras):
                            palabra = palabras[i]
                            if len(palabra) == 1 and palabra.isalpha():
                                palabra_completa = palabra
                                j = i + 1
                                while j < len(palabras) and len(palabras[j]) == 1 and palabras[j].isalpha():
                                    palabra_completa += palabras[j]
                                    j += 1
                                texto_reconstruido += palabra_completa + " "
                                i = j
                            else:
                                texto_reconstruido += palabra + " "
                                i += 1
                        
                        texto_reconstruido = texto_reconstruido.strip()
                        
                        # Buscar VALOR A PAGAR
                        if not valor_a_pagar:
                            patron_valor = r'VALORAPAGARACOMERCIO[\s\$]*([\d,\.]+)'
                            match = re.search(patron_valor, texto_reconstruido, re.IGNORECASE)
                            if match:
                                valor_texto = match.group(1)
                                valor_limpio = valor_texto.replace(',', '').replace('.', '')
                                if valor_limpio.isdigit():
                                    valor_num = int(valor_limpio)
                                    if valor_num > 1000000:
                                        valor_a_pagar = valor_num
                                        st.success(f"💰 Valor encontrado: ${valor_a_pagar:,.0f}")
                        
                        # Buscar CANTIDAD PASOS
                        if not cantidad_pasos:
                            patron_pasos = r'CANTIDADPASOS[\s]*([\d,\.]+)'
                            match = re.search(patron_pasos, texto_reconstruido, re.IGNORECASE)
                            if match:
                                pasos_texto = match.group(1)
                                pasos_limpio = pasos_texto.replace(',', '').replace('.', '')
                                if pasos_limpio.isdigit():
                                    pasos_num = int(pasos_limpio)
                                    if 1000 <= pasos_num <= 100000:
                                        cantidad_pasos = pasos_num
                                        st.success(f"👣 Pasos encontrados: {cantidad_pasos:,}")
                        
                        if valor_a_pagar and cantidad_pasos:
                            break
        
        return valor_a_pagar, cantidad_pasos
            
    except Exception as e:
        st.error(f"❌ Error extrayendo datos: {str(e)}")
        return None, None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiNzU2ZTI0NWEtNjIxOC00NmMzLThiODItNjk2YmNhM2QyMjMwIiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None, None
    
    try:
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None, None
        
        time.sleep(6)
        
        with st.spinner("📊 Extrayendo datos de ACCENORTE..."):
            valor_power_bi, pasos_power_bi = find_accenorte_data(driver)
        
        return valor_power_bi, pasos_power_bi
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None, None
    finally:
        if driver:
            driver.quit()

def comparar_valores(valor_excel, valor_power_bi, pasos_excel, pasos_power_bi):
    """Compara los valores y determina si coinciden - VALIDACIÓN ESTRICTA"""
    try:
        if valor_power_bi is None or pasos_power_bi is None:
            return False, False, 0, 0
            
        diferencia_valor = abs(valor_excel - valor_power_bi)
        diferencia_pasos = abs(pasos_excel - pasos_power_bi)
        
        # VALIDACIÓN ESTRICTA - DEBEN COINCIDIR EXACTAMENTE
        coinciden_valor = diferencia_valor == 0
        coinciden_pasos = diferencia_pasos == 0
        
        return coinciden_valor, coinciden_pasos, diferencia_valor, diferencia_pasos
        
    except Exception as e:
        return False, False, 0, 0

# ===== INTERFAZ PRINCIPAL =====

def main():
    
    # Sidebar
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 20px 0;">
            <h2>📋 Validador ACCENORTE</h2>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.info("""
        **Instrucciones:**
        1. Cargar archivo Excel
        2. Verificar fecha detectada
        3. Validar automáticamente con Power BI
        4. Revisar resultados
        """)
        
        st.markdown("---")
        
        st.success("**Estado:** ✅ Sistema Operativo")
        st.info(f"**Python:** {sys.version_info.major}.{sys.version_info.minor}")
        st.info(f"**Pandas:** {pd.__version__}")
    
    # Cargar archivo Excel
    st.markdown("### 📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel de ACCENORTE", 
        type=['xlsx', 'xls'],
        label_visibility="collapsed"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Archivo cargado: {uploaded_file.name}")
        
        fecha_validacion = extraer_fecha_desde_excel(uploaded_file)
        
        if not fecha_validacion:
            st.warning("⚠️ No se pudo detectar la fecha automáticamente")
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):")
        
        if fecha_validacion:
            with st.spinner("📊 Procesando archivo Excel..."):
                valor_excel, pasos_excel = procesar_excel(uploaded_file)
            
            if valor_excel > 0 and pasos_excel > 0:
                # Mostrar valores del Excel
                st.markdown("### 📊 Valores Extraídos del Excel")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("""
                    <div class="metric-card">
                        <h3>💰 Valor a Pagar</h3>
                        <h2>${:,.0f}</h2>
                    </div>
                    """.format(valor_excel), unsafe_allow_html=True)
                with col2:
                    st.markdown("""
                    <div class="metric-card">
                        <h3>👣 Número de Pasos</h3>
                        <h2>{:,}</h2>
                    </div>
                    """.format(pasos_excel), unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Extracción de Power BI
                st.markdown("### 🌐 Extracción de Power BI")
                
                with st.spinner("Conectando y extrayendo datos del Power BI..."):
                    valor_power_bi, pasos_power_bi = extract_powerbi_data(fecha_validacion)
                    
                    if valor_power_bi is not None and pasos_power_bi is not None:
                        # Mostrar resultados de Power BI
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col3, col4 = st.columns(2)
                        with col3:
                            st.markdown("""
                            <div class="metric-card">
                                <h3>💰 VALOR A PAGAR A COMERCIO</h3>
                                <h2>${:,.0f}</h2>
                            </div>
                            """.format(valor_power_bi), unsafe_allow_html=True)
                        with col4:
                            st.markdown("""
                            <div class="metric-card">
                                <h3>👣 CANTIDAD PASOS</h3>
                                <h2>{:,}</h2>
                            </div>
                            """.format(pasos_power_bi), unsafe_allow_html=True)
                        
                        st.markdown("---")
                        
                        # Comparación ESTRICTA
                        st.markdown("### 📊 Resultado de la Validación")
                        
                        coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                            valor_excel, valor_power_bi, pasos_excel, pasos_power_bi
                        )
                        
                        # Mostrar resultado principal
                        if coinciden_valor and coinciden_pasos:
                            st.markdown("""
                            <div class="success-box">
                                🎉 ✅ TODOS LOS VALORES COINCIDEN EXACTAMENTE
                            </div>
                            """, unsafe_allow_html=True)
                            st.balloons()
                        else:
                            st.markdown("""
                            <div class="error-box">
                                ❌ LOS VALORES NO COINCIDEN
                            </div>
                            """, unsafe_allow_html=True)
                            
                            if not coinciden_valor:
                                st.error(f"**Diferencia en VALOR:** ${dif_valor:,.0f}")
                                col5, col6 = st.columns(2)
                                with col5:
                                    st.error(f"**Excel:** ${valor_excel:,.0f}")
                                with col6:
                                    st.error(f"**Power BI:** ${valor_power_bi:,.0f}")
                            
                            if not coinciden_pasos:
                                st.error(f"**Diferencia en PASOS:** {dif_pasos} pasos")
                                col7, col8 = st.columns(2)
                                with col7:
                                    st.error(f"**Excel:** {pasos_excel}")
                                with col8:
                                    st.error(f"**Power BI:** {pasos_power_bi:,}")
                        
                        # Tabla resumen
                        st.markdown("### 📋 Resumen de Comparación")
                        
                        datos = {
                            'Concepto': ['Valor a Pagar', 'Número de Pasos'],
                            'Excel': [f"${valor_excel:,.0f}", f"{pasos_excel}"],
                            'Power BI': [f"${valor_power_bi:,.0f}", f"{pasos_power_bi:,}"],
                            'Resultado': [
                                '✅ COINCIDE' if coinciden_valor else f'❌ DIFERENCIA: ${dif_valor:,.0f}',
                                '✅ COINCIDE' if coinciden_pasos else f'❌ DIFERENCIA: {dif_pasos} pasos'
                            ]
                        }
                        
                        df = pd.DataFrame(datos)
                        st.dataframe(df, use_container_width=True, hide_index=True)
                        
                    else:
                        st.markdown("""
                        <div class="error-box">
                            ❌ No se pudieron extraer los datos de Power BI
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="error-box">
                    ❌ No se pudieron extraer los valores del Excel
                </div>
                """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="text-align: center; padding: 50px 20px; color: #666;">
            <h3>📁 Por favor, carga un archivo Excel para comenzar</h3>
            <p>Selecciona un archivo Excel de ACCENORTE en el campo superior</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
    
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #888; padding: 20px 0;">
        <p>💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v5.0</p>
    </div>
    """, unsafe_allow_html=True)
