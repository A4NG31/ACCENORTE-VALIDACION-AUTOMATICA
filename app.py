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

.stSpinner > div > div {
    border-color: #00CFFF !important;
}

.stProgress > div > div > div > div {
    background-color: #00CFFF !important;
}

.success-box {
    background-color: #d4edda;
    border: 1px solid #c3e6cb;
    border-radius: 5px;
    padding: 15px;
    margin: 10px 0;
    color: #155724;
}

.error-box {
    background-color: #f8d7da;
    border: 1px solid #f5c6cb;
    border-radius: 5px;
    padding: 15px;
    margin: 10px 0;
    color: #721c24;
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

# ===== FUNCIONES MEJORADAS =====

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
                            return fecha.strftime("%Y-%m-%d")
        
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
                        texto = elemento.text.strip()
                        if 'ACCENORTE' in texto.upper() and fecha_objetivo in texto:
                            elemento_conciliacion = elemento
                            break
                if elemento_conciliacion:
                    break
            except:
                continue
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_conciliacion)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(3)
            return True
        else:
            return False
            
    except Exception as e:
        return False

def find_accenorte_data(driver):
    """Extrae los valores de VALOR A PAGAR A COMERCIO y CANTIDAD PASOS"""
    try:
        valor_a_pagar = None
        cantidad_pasos = None
        
        # Buscar en elementos de la esquina superior izquierda
        elementos_esquina = driver.find_elements(By.XPATH, "//*[position() < 20]")
        
        for elemento in elementos_esquina:
            if elemento.is_displayed():
                location = elemento.location
                if location['x'] < 600 and location['y'] < 600:
                    texto_completo = elemento.text.strip()
                    if texto_completo:
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
                        
                        if valor_a_pagar and cantidad_pasos:
                            break
        
        return valor_a_pagar, cantidad_pasos
            
    except Exception as e:
        return None, None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiNzU2ZTI0NWEtNjIxOC00NmMzLThiODItNjk2YmNhM2QyMjMwIiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None, None
    
    try:
        with st.spinner("Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(8)
        
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None, None
        
        time.sleep(5)
        
        with st.spinner("Extrayendo datos..."):
            valor_power_bi, pasos_power_bi = find_accenorte_data(driver)
        
        return valor_power_bi, pasos_power_bi
        
    except Exception as e:
        return None, None
    finally:
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
    st.title("💰 Validador Power BI - ACCENORTE")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("📋 Información")
    st.sidebar.info("""
    **Objetivo:**
    - Validar conciliaciones entre Excel y Power BI
    - Comparar valores exactos
    - Verificar número de pasos
    """)
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel de ACCENORTE", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        fecha_validacion = extraer_fecha_desde_excel(uploaded_file)
        
        if not fecha_validacion:
            st.warning("No se pudo detectar la fecha en el rango G18:N24")
            fecha_validacion = st.text_input("Ingresa la fecha manualmente (YYYY-MM-DD):")
        
        if fecha_validacion:
            with st.spinner("Procesando archivo Excel..."):
                valor_excel, pasos_excel = procesar_excel(uploaded_file)
            
            if valor_excel > 0 and pasos_excel > 0:
                # Mostrar valores del Excel
                st.markdown("### 📊 Valores del Excel")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("💰 Valor a Pagar", f"${valor_excel:,.0f}")
                with col2:
                    st.metric("👣 Número de Pasos", f"{pasos_excel}")
                
                st.markdown("---")
                
                # Extracción de Power BI
                with st.spinner("Conectando con Power BI..."):
                    valor_power_bi, pasos_power_bi = extract_powerbi_data(fecha_validacion)
                    
                    if valor_power_bi is not None and pasos_power_bi is not None:
                        # Mostrar resultados de Power BI
                        st.markdown("### 📊 Valores de Power BI")
                        
                        col3, col4 = st.columns(2)
                        with col3:
                            st.metric("💰 VALOR A PAGAR A COMERCIO", f"${valor_power_bi:,.0f}")
                        with col4:
                            st.metric("👣 CANTIDAD PASOS", f"{pasos_power_bi:,}")
                        
                        st.markdown("---")
                        
                        # Comparación ESTRICTA
                        st.markdown("### 📊 Resultado de la Validación")
                        
                        coinciden_valor, coinciden_pasos, dif_valor, dif_pasos = comparar_valores(
                            valor_excel, valor_power_bi, pasos_excel, pasos_power_bi
                        )
                        
                        # Mostrar resultado principal
                        if coinciden_valor and coinciden_pasos:
                            st.success("🎉 ✅ TODOS LOS VALORES COINCIDEN EXACTAMENTE")
                            st.balloons()
                        else:
                            st.error("❌ LOS VALORES NO COINCIDEN")
                            
                            if not coinciden_valor:
                                st.error(f"**Diferencia en VALOR:** ${dif_valor:,.0f}")
                                st.error(f"Excel: ${valor_excel:,.0f} vs Power BI: ${valor_power_bi:,.0f}")
                            
                            if not coinciden_pasos:
                                st.error(f"**Diferencia en PASOS:** {dif_pasos} pasos")
                                st.error(f"Excel: {pasos_excel} vs Power BI: {pasos_power_bi:,}")
                        
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
                        st.error("No se pudieron extraer los datos de Power BI")
            else:
                st.error("No se pudieron extraer los valores del Excel")
    else:
        st.info("Por favor, carga un archivo Excel para comenzar")

if __name__ == "__main__":
    main()
    
    st.markdown("---")
    st.markdown('<div style="text-align: center;">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit</div>', unsafe_allow_html=True)
