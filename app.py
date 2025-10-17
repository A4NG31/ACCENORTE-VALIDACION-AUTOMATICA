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

# ===== CSS =====
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
    width: 300px !important;
    padding: 20px 10px 20px 10px !important;
    border-right: 1px solid #333 !important;
}

[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}

[data-testid="stSidebar"] .stFileUploader > label {
    color: white !important;
    font-weight: bold;
}

[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-title,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-subtitle,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-list button,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-name,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-status,
[data-testid="stSidebar"] .stFileUploader span,
[data-testid="stSidebar"] .stFileUploader div {
    color: black !important;
}

[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button {
    color: black !important;
    background-color: #f0f0f0 !important;
    border: 1px solid #ccc !important;
}

[data-testid="stSidebar"] svg.icon {
    stroke: white !important;
    color: white !important;
    fill: none !important;
    opacity: 1 !important;
}

.stSpinner > div > div {
    border-color: #00CFFF !important;
}

.stProgress > div > div > div > div {
    background-color: #00CFFF !important;
}
</style>
""", unsafe_allow_html=True)

# Logo
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES MODIFICADAS =====

def extraer_fecha_desde_excel(uploaded_file):
    """Extrae la fecha desde la celda combinada (G18:N24) del Excel"""
    try:
        df = pd.read_excel(uploaded_file, header=None)
        
        # Buscar en el rango G18:N24 (índices 6:13 para columnas, 17:23 para filas)
        for fila in range(17, 24):  # Filas 18-24 (0-indexed: 17-23)
            for col in range(6, 14):  # Columnas G-N (0-indexed: 6-13)
                if pd.notna(df.iloc[fila, col]):
                    celda = str(df.iloc[fila, col]).strip()
                    # Buscar patrones de fecha
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
                            st.success(f"📅 Fecha encontrada en Excel: {fecha.strftime('%d/%m/%Y')}")
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
                # Sumar valores debajo del encabezado
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
    """Hacer clic en la conciliación específica por fecha - ACCENORTE"""
    try:
        # Formatear fecha para búsqueda
        fecha_formateada = f"{fecha_objetivo} 00:00 al {fecha_objetivo} 11:59"
        
        # Buscar el elemento que contiene la fecha exacta
        selectors = [
            f"//*[contains(text(), 'Conciliación Accenorte del {fecha_formateada}')]",
            f"//*[contains(text(), 'CONCILIACIÓN ACCENORTE DEL {fecha_formateada}')]",
            f"//*[contains(text(), '{fecha_formateada}')]",
            f"//div[contains(text(), '{fecha_objetivo}')]",
            f"//span[contains(text(), '{fecha_objetivo}')]",
        ]
        
        elemento_conciliacion = None
        for selector in selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if 'ACCENORTE' in texto.upper() and fecha_objetivo in texto:
                            elemento_conciliacion = elemento
                            st.success(f"✅ Encontrado: {elemento.text.strip()}")
                            break
                if elemento_conciliacion:
                    break
            except:
                continue
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView(true);", elemento_conciliacion)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(3)
            return True
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            return False
            
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliación: {str(e)}")
        return False

def find_accenorte_data(driver):
    """
    Buscar los valores de VALOR A PAGAR A COMERCIO y CANTIDAD PASOS
    """
    try:
        valor_a_pagar = None
        cantidad_pasos = None
        
        # ESTRATEGIA 1: Buscar por títulos específicos
        titulos_buscar = [
            "VALOR A PAGAR A COMERCIO",
            "CANTIDAD PASOS"
        ]
        
        for titulo in titulos_buscar:
            try:
                # Buscar el título
                titulo_element = None
                selectors = [
                    f"//*[contains(text(), '{titulo}')]",
                    f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{titulo.lower()}')]"
                ]
                
                for selector in selectors:
                    try:
                        elementos = driver.find_elements(By.XPATH, selector)
                        for elemento in elementos:
                            if elemento.is_displayed():
                                titulo_element = elemento
                                break
                        if titulo_element:
                            break
                    except:
                        continue
                
                if titulo_element:
                    # Buscar el valor asociado al título
                    # Estrategia: buscar en el mismo contenedor o elemento hermano
                    parent = titulo_element.find_element(By.XPATH, "./..")
                    all_text = parent.text
                    
                    # Extraer números del texto
                    if "VALOR A PAGAR" in titulo.upper():
                        # Buscar formato monetario: $102.031.300
                        valor_match = re.search(r'\$[\d\.]+(?:\.\d{3})*', all_text)
                        if valor_match:
                            valor_texto = valor_match.group(0)
                            # Limpiar: $102.031.300 -> 102031300
                            valor_limpio = valor_texto.replace('$', '').replace('.', '')
                            if valor_limpio.isdigit():
                                valor_a_pagar = int(valor_limpio)
                                st.success(f"💰 VALOR A PAGAR: ${valor_a_pagar:,.0f}")
                    
                    elif "CANTIDAD PASOS" in titulo.upper():
                        # Buscar formato: 6.704
                        pasos_match = re.search(r'\b\d{1,3}(?:\.\d{3})*\b', all_text)
                        if pasos_match:
                            pasos_texto = pasos_match.group(0)
                            # Limpiar: 6.704 -> 6704
                            pasos_limpio = pasos_texto.replace('.', '')
                            if pasos_limpio.isdigit():
                                cantidad_pasos = int(pasos_limpio)
                                st.success(f"👣 CANTIDAD PASOS: {cantidad_pasos}")
            
            except Exception as e:
                st.warning(f"⚠️ Error buscando {titulo}: {e}")
        
        # ESTRATEGIA 2: Buscar en tablas o cards
        if not valor_a_pagar or not cantidad_pasos:
            try:
                # Buscar todos los elementos que contengan números grandes
                elementos_numeros = driver.find_elements(By.XPATH, "//*[contains(text(), '$') or contains(text(), '.')]")
                
                for elemento in elementos_numeros:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        
                        # Buscar valor con formato $XXX.XXX.XXX
                        if not valor_a_pagar and '$' in texto:
                            valor_match = re.search(r'\$[\d\.]+(?:\.\d{3})*', texto)
                            if valor_match:
                                valor_texto = valor_match.group(0)
                                valor_limpio = valor_texto.replace('$', '').replace('.', '')
                                if valor_limpio.isdigit() and len(valor_limpio) >= 6:  # Valores grandes
                                    valor_a_pagar = int(valor_limpio)
                                    st.success(f"💰 Valor encontrado (estrategia 2): ${valor_a_pagar:,.0f}")
                        
                        # Buscar pasos con formato X.XXX
                        if not cantidad_pasos and re.search(r'\b\d{1,3}\.\d{3}\b', texto):
                            pasos_match = re.search(r'\b\d{1,3}\.\d{3}\b', texto)
                            if pasos_match:
                                pasos_texto = pasos_match.group(0)
                                pasos_limpio = pasos_texto.replace('.', '')
                                if pasos_limpio.isdigit() and 1000 <= int(pasos_limpio) <= 99999:
                                    cantidad_pasos = int(pasos_limpio)
                                    st.success(f"👣 Pasos encontrados (estrategia 2): {cantidad_pasos}")
            
            except Exception as e:
                st.warning(f"⚠️ Estrategia 2 falló: {e}")
        
        # Validar resultados
        if valor_a_pagar and cantidad_pasos:
            st.success(f"✅ Extracción exitosa: Valor=${valor_a_pagar:,.0f}, Pasos={cantidad_pasos}")
            return valor_a_pagar, cantidad_pasos
        else:
            st.error(f"❌ Extracción parcial: Valor={valor_a_pagar}, Pasos={cantidad_pasos}")
            return valor_a_pagar, cantidad_pasos
            
    except Exception as e:
        st.error(f"❌ Error buscando datos ACCENORTE: {str(e)}")
        return None, None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - ACCENORTE"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiNzU2ZTI0NWEtNjIxOC00NmMzLThiODItNjk2YmNhM2QyMjMwIiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None, None
    
    try:
        # 1. Navegar al reporte
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        # 2. Tomar screenshot inicial
        driver.save_screenshot("powerbi_inicial.png")
        
        # 3. Hacer clic en la conciliación específica
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None, None
        
        # 4. Esperar a que cargue la selección
        time.sleep(5)
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        # 5. Buscar datos de ACCENORTE
        with st.spinner("🔍 Extrayendo datos de ACCENORTE..."):
            valor_power_bi, pasos_power_bi = find_accenorte_data(driver)
        
        # 6. Tomar screenshot final
        driver.save_screenshot("powerbi_final.png")
        
        return valor_power_bi, pasos_power_bi
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None, None
    finally:
        driver.quit()

def comparar_valores(valor_excel, valor_power_bi, pasos_excel, pasos_power_bi):
    """Compara los valores y determina si coinciden"""
    try:
        diferencia_valor = abs(valor_excel - valor_power_bi) if valor_power_bi else valor_excel
        diferencia_pasos = abs(pasos_excel - pasos_power_bi) if pasos_power_bi else pasos_excel
        
        coinciden_valor = diferencia_valor < 1.0
        coinciden_pasos = diferencia_pasos == 0
        
        return coinciden_valor, coinciden_pasos, diferencia_valor, diferencia_pasos
        
    except Exception as e:
        st.error(f"❌ Error comparando valores: {e}")
        return False, False, 0, 0

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.title("💰 Validador Power BI - ACCENORTE")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Validar conciliaciones entre Excel y Power BI
    - Extraer datos de ACCENORTE
    - Comparar valores y número de pasos
    
    **Formato Excel:**
    - Fecha en celda combinada G18:N24
    - Valor en columna AK
    - Total transacciones en texto
    
    **Power BI:**
    - Conciliación Accenorte del YYYY-MM-DD
    - VALOR A PAGAR A COMERCIO ($)
    - CANTIDAD PASOS (formato X.XXX)
    """)
    
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel de ACCENORTE", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Extraer fecha del Excel
        fecha_validacion = extraer_fecha_desde_excel(uploaded_file)
        
        if not fecha_validacion:
            st.warning("⚠️ No se pudo detectar la fecha en el rango G18:N24")
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):")
        
        if fecha_validacion:
            # Procesar Excel
            with st.spinner("📊 Procesando archivo Excel..."):
                valor_excel, pasos_excel = procesar_excel(uploaded_file)
            
            if valor_excel > 0 and pasos_excel > 0:
                # Mostrar valores del Excel
                st.markdown("### 📊 Valores Extraídos del Excel")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("💰 Valor a Pagar", f"${valor_excel:,.0f}")
                with col2:
                    st.metric("👣 Número de Pasos", f"{pasos_excel}")
                
                st.markdown("---")
                
                # EXTRACCIÓN AUTOMÁTICA
                st.info(f"🤖 **Extracción Automática Activada** - Buscando conciliación del {fecha_validacion}...")
                
                with st.spinner("🌐 Extrayendo datos de Power BI..."):
                    valor_power_bi, pasos_power_bi = extract_powerbi_data(fecha_validacion)
                    
                    if valor_power_bi is not None and pasos_power_bi is not None:
                        # Mostrar resultados de Power BI
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col3, col4 = st.columns(2)
                        with col3:
                            st.metric("💰 VALOR A PAGAR A COMERCIO", f"${valor_power_bi:,.0f}")
                        with col4:
                            st.metric("👣 CANTIDAD PASOS", f"{pasos_power_bi:,}")
                        
                        st.markdown("---")
                        
                        # Comparar
                        st.markdown("### 📊 Resultado de la Validación")
                        
                        coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                            valor_excel, valor_power_bi, pasos_excel, pasos_power_bi
                        )
                        
                        if coinciden_valor and coinciden_pasos:
                            st.success("🎉 ✅ TODOS LOS VALORES COINCIDEN")
                            st.balloons()
                        else:
                            if not coinciden_valor:
                                st.error(f"❌ DIFERENCIA EN VALOR: ${dif_valor:,.0f}")
                            if not coinciden_pasos:
                                st.error(f"❌ DIFERENCIA EN PASOS: {dif_pasos} pasos")
                        
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
                        
                        # Screenshots
                        with st.expander("📸 Ver Capturas del Proceso"):
                            col1, col2, col3 = st.columns(3)
                            
                            if os.path.exists("powerbi_inicial.png"):
                                with col1:
                                    st.image("powerbi_inicial.png", caption="Vista Inicial", use_column_width=True)
                            
                            if os.path.exists("powerbi_despues_seleccion.png"):
                                with col2:
                                    st.image("powerbi_despues_seleccion.png", caption="Tras Selección", use_column_width=True)
                            
                            if os.path.exists("powerbi_final.png"):
                                with col3:
                                    st.image("powerbi_final.png", caption="Vista Final", use_column_width=True)
                    else:
                        st.error("❌ No se pudieron extraer los datos de Power BI")
            else:
                st.error("❌ No se pudieron extraer los valores del Excel")
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar")
    
    # Ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. Cargar archivo Excel de ACCENORTE
        2. Extracción automática de fecha (rango G18:N24)
        3. Extracción de valores del Excel (columna AK y TOTAL TRANSACCIONES)
        4. Conexión con Power BI y selección de fecha
        5. Extracción de datos: VALOR A PAGAR A COMERCIO y CANTIDAD PASOS
        6. Comparación y validación
        
        **Formato esperado en Power BI:**
        - Título: `Conciliación Accenorte del YYYY-MM-DD 00:00 al YYYY-MM-DD 11:59`
        - Valor: `VALOR A PAGAR A COMERCIO` con formato `$102.031.300`
        - Pasos: `CANTIDAD PASOS` con formato `6.704`
        
        **Notas:**
        - La fecha se busca en celdas combinadas G18:N24 del Excel
        - El valor se suma de la columna AK debajo del encabezado "Valor"
        - Los pasos se buscan en "TOTAL TRANSACCIONES X"
        """)

if __name__ == "__main__":
    main()
    
    st.markdown("---")
    st.markdown('<div style="text-align: center;">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v4.0 - ACCENORTE</div>', unsafe_allow_html=True)