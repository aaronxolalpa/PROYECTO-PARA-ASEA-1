from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import pandas as pd

# Ruta del archivo Excel con las guías y periodos
archivo_excel = 'C:/Users/sergio/Desktop/pantallas_sepomex/guias_leer.xlsx'

def obtener_guias_y_periodos(archivo):
    df = pd.read_excel(archivo, dtype=str)  # Leer el archivo Excel
    return df.values.tolist()  # Convertir a lista de listas

# Leer las guías y periodos del archivo
lista_guias = obtener_guias_y_periodos(archivo_excel)

# Configurar la carpeta de destino para guardar el PDF (Carpeta "busquedas")
descargas_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'busquedas')
if not os.path.exists(descargas_dir):
    os.makedirs(descargas_dir)

def configurar_navegador():
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "savefile.default_directory": descargas_dir,
        "savefile.prompt_for_download": False,
        "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
        "printing.default_destination_selection_rules": '{"kind":"local","name":"Save as PDF"}'
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--kiosk-printing')  # Imprime automáticamente sin mostrar el diálogo
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def procesar_guia(guia_numero, periodo):
    driver = configurar_navegador()
    url = 'https://www.correosdemexico.gob.mx/SSLServicios/SeguimientoEnvio/Seguimiento.aspx'
    try:
        driver.get(url)
        driver.implicitly_wait(10)
        driver.find_element(By.NAME, 'Guia').send_keys(guia_numero)
        driver.find_element(By.NAME, 'Periodo').send_keys(periodo)
        driver.find_element(By.NAME, 'Busqueda').click()
        time.sleep(10)
        driver.execute_script('window.print();')
        time.sleep(10)
        renombrar_pdf(descargas_dir, guia_numero)
    except Exception as e:
        print(f"Error procesando guía {guia_numero}: {e}")
    finally:
        driver.quit()

def renombrar_pdf(descargas_dir, guia_numero):
    time.sleep(5)
    for archivo in os.listdir(descargas_dir):
        if archivo.endswith(".pdf") and "Seguimiento" in archivo:
            origen = os.path.join(descargas_dir, archivo)
            destino = os.path.join(descargas_dir, f"{guia_numero}.pdf")
            os.rename(origen, destino)
            break

for guia, periodo in lista_guias:
    procesar_guia(guia, periodo)

print("Todos los resultados de búsqueda se han guardado correctamente.")
