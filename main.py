import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import pandas as pd
import threading
import subprocess
import platform


def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        entrada_archivo.delete(0, tk.END)
        entrada_archivo.insert(0, archivo)


def obtener_guias_y_periodos(archivo):
    df = pd.read_excel(archivo, dtype=str)
    return df.values.tolist()


def configurar_navegador(descargas_dir):
    chrome_options = webdriver.ChromeOptions()
    
    # Configurar la carpeta de descargas directamente a la carpeta 'busquedas'
    prefs = {
        "download.default_directory": descargas_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
        "printing.default_destination_selection_rules": '{"kind":"local","name":"Save as PDF"}',
        "savefile.default_directory": descargas_dir,
        "savefile.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--kiosk-printing')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-plugins')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)


def esperar_descarga_y_renombrar(descargas_dir, guia_numero, timeout=30):
    """Espera a que se complete la descarga y renombra el archivo"""
    tiempo_inicial = time.time()
    archivo_temporal = None
    
    while time.time() - tiempo_inicial < timeout:
        archivos = os.listdir(descargas_dir)
        
        # Buscar archivos .crdownload (descarga en progreso)
        archivos_descargando = [f for f in archivos if f.endswith('.crdownload')]
        if archivos_descargando:
            time.sleep(1)
            continue
        
        # Buscar el archivo PDF recién descargado
        archivos_pdf = [f for f in archivos if f.endswith('.pdf')]
        
        # Buscar archivos que contengan palabras clave del seguimiento
        for archivo in archivos_pdf:
            if any(palabra in archivo.lower() for palabra in ['seguimiento', 'tracking', 'sepomex', 'correos']):
                # Verificar que no sea uno de nuestros archivos ya renombrados
                if not archivo.replace('.pdf', '').isdigit():
                    archivo_temporal = archivo
                    break
        
        if archivo_temporal:
            break
        
        time.sleep(1)
    
    # Renombrar el archivo si se encontró
    if archivo_temporal:
        try:
            origen = os.path.join(descargas_dir, archivo_temporal)
            destino = os.path.join(descargas_dir, f"{guia_numero}.pdf")
            
            # Si ya existe un archivo con ese nombre, eliminarlo
            if os.path.exists(destino):
                os.remove(destino)
            
            os.rename(origen, destino)
            print(f"Archivo renombrado: {archivo_temporal} -> {guia_numero}.pdf")
            return True
        except Exception as e:
            print(f"Error al renombrar archivo para guía {guia_numero}: {e}")
            return False
    else:
        print(f"No se pudo encontrar el archivo descargado para la guía {guia_numero}")
        return False


def procesar_guia(guia_numero, periodo, descargas_dir):
    driver = configurar_navegador(descargas_dir)
    url = 'https://www.correosdemexico.gob.mx/SSLServicios/SeguimientoEnvio/Seguimiento.aspx'
    
    try:
        driver.get(url)
        driver.implicitly_wait(2)
        
        # Llenar el formulario
        driver.find_element(By.NAME, 'Guia').send_keys(str(guia_numero))
        driver.find_element(By.NAME, 'Periodo').send_keys(str(periodo))
        driver.find_element(By.NAME, 'Busqueda').click()
        
        # Esperar a que cargue la página de resultados
        time.sleep(2)
        
        # Imprimir a PDF (esto debería descargar directamente a la carpeta configurada)
        driver.execute_script('window.print();')
        
        # Esperar a que se complete la descarga y renombrar
        esperar_descarga_y_renombrar(descargas_dir, guia_numero)
        
    except Exception as e:
        print(f"Error procesando guía {guia_numero}: {e}")
    finally:
        driver.quit()


def abrir_carpeta_busquedas(ruta):
    sistema = platform.system()
    if sistema == "Windows":
        os.startfile(ruta)
    elif sistema == "Darwin":
        subprocess.Popen(["open", ruta])
    else:
        subprocess.Popen(["xdg-open", ruta])


def iniciar_proceso():
    archivo_excel = entrada_archivo.get()
    if not archivo_excel:
        messagebox.showerror("Error", "Selecciona un archivo Excel.")
        return

    boton_iniciar.config(state=tk.DISABLED)
    boton_buscar.config(state=tk.DISABLED)
    progreso['value'] = 0
    estado.set("Procesando...")

    threading.Thread(target=procesar_guias, args=(archivo_excel,), daemon=True).start()


def procesar_guias(archivo_excel):
    try:
        lista_guias = obtener_guias_y_periodos(archivo_excel)
        total_guias = len(lista_guias)
        
        # Crear la carpeta 'busquedas' en el mismo directorio que el script
        descargas_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'busquedas')
        os.makedirs(descargas_dir, exist_ok=True)
        
        print(f"Los archivos se guardarán en: {descargas_dir}")

        for index, (guia, periodo) in enumerate(lista_guias, 1):
            estado.set(f"Procesando guía {index}/{total_guias}: {guia}")
            ventana.update_idletasks()
            
            procesar_guia(guia, periodo, descargas_dir)
            
            progreso['value'] = (index / total_guias) * 100
            ventana.update_idletasks()
            
            # Pequeña pausa entre descargas para evitar sobrecargar el servidor
            time.sleep(1)

        estado.set("Proceso completado.")
        messagebox.showinfo("Completado", f"Todos los resultados de búsqueda se han guardado en:\n{descargas_dir}")
        abrir_carpeta_busquedas(descargas_dir)
        
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        estado.set("Error en el proceso.")
    finally:
        boton_iniciar.config(state=tk.NORMAL)
        boton_buscar.config(state=tk.NORMAL)


# Interfaz gráfica mejorada
ventana = tk.Tk()
ventana.title("Descarga de Seguimientos Sepomex")
ventana.geometry("550x250")
ventana.resizable(False, False)

frame = tk.Frame(ventana, padx=20, pady=20)
frame.pack(expand=True)

tk.Label(frame, text="Selecciona el archivo Excel:", font=("Arial", 12)).grid(row=0, column=0, sticky="w")

entrada_archivo = tk.Entry(frame, width=45, font=("Arial", 10))
entrada_archivo.grid(row=1, column=0, padx=(0, 10), pady=5)

boton_buscar = tk.Button(frame, text="Buscar", command=seleccionar_archivo, width=10)
boton_buscar.grid(row=1, column=1)

boton_iniciar = tk.Button(frame, text="Iniciar Proceso", command=iniciar_proceso, bg="green", fg="white", font=("Arial", 11), width=20)
boton_iniciar.grid(row=2, column=0, columnspan=2, pady=20)

progreso = ttk.Progressbar(frame, orient="horizontal", length=400, mode="determinate")
progreso.grid(row=3, column=0, columnspan=2, pady=10)

estado = tk.StringVar()
estado.set("Esperando...")
tk.Label(frame, textvariable=estado, font=("Arial", 10, "italic")).grid(row=4, column=0, columnspan=2)

ventana.mainloop()