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
import shutil
import sys



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
    prefs = {
        "savefile.default_directory": descargas_dir,
        "savefile.prompt_for_download": False,
        "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
        "printing.default_destination_selection_rules": '{"kind":"local","name":"Save as PDF"}'
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--kiosk-printing')
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)


def renombrar_pdf(descargas_dir, guia_numero):
    time.sleep(5)
    for archivo in os.listdir(descargas_dir):
        if archivo.endswith(".pdf") and "Seguimiento" in archivo:
            origen = os.path.join(descargas_dir, archivo)
            destino = os.path.join(descargas_dir, f"{guia_numero}.pdf")
            os.rename(origen, destino)
            break


def procesar_guia(guia_numero, periodo, descargas_dir):
    driver = configurar_navegador(descargas_dir)
    url = 'https://www.correosdemexico.gob.mx/SSLServicios/SeguimientoEnvio/Seguimiento.aspx'
    try:
        driver.get(url)
        driver.implicitly_wait(1)
        driver.find_element(By.NAME, 'Guia').send_keys(guia_numero)
        driver.find_element(By.NAME, 'Periodo').send_keys(periodo)
        driver.find_element(By.NAME, 'Busqueda').click()
        time.sleep(1)
        driver.execute_script('window.print();')
        renombrar_pdf(descargas_dir, guia_numero)
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

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        # Si está congelado (ejecutable .exe)
        return os.path.dirname(sys.executable)
    else:
        # Si es un script .py
        return os.path.dirname(os.path.realpath(__file__))


def procesar_guias(archivo_excel):
    lista_guias = obtener_guias_y_periodos(archivo_excel)
    total_guias = len(lista_guias)

    # Obtener ruta base del .py o .exe
    ruta_base = obtener_ruta_base()

    #Crear carpeta  donde se guardarán los PDFs
    carpeta = os.path.join(ruta_base, 'Busquedas')
    if os.path.exists(carpeta):
        shutil.rmtree(carpeta)
    os.makedirs(carpeta)

    # Procesar cada guía
    for index, (guia, periodo) in enumerate(lista_guias, 1):
        procesar_guia(guia, periodo, carpeta)
        progreso['value'] = (index / total_guias) * 100
        estado.set(f"Procesando guía {index}/{total_guias}")
        ventana.update_idletasks()

    # Mostrar mensaje y abrir carpeta de resultados
    estado.set("Proceso completado.")
    messagebox.showinfo("Completado", "Todos los resultados se copiaron correctamente.")
    abrir_carpeta_busquedas(carpeta)

    #Reactivar botones
    boton_iniciar.config(state=tk.NORMAL)
    boton_buscar.config(state=tk.NORMAL)
def verificar_busquedas_al_inicio():
    ruta_base = obtener_ruta_base()
    carpeta = os.path.join(ruta_base, 'Busquedas')

    if os.path.exists(carpeta):
        archivos = os.listdir(carpeta)
        if archivos:
            respuesta = messagebox.askyesno(
                "Archivos encontrados",
                f"Se encontraron {len(archivos)} archivo(s) en la carpeta 'Busquedas'.\n¿Deseas eliminarlos?"
            )
            if respuesta:
                try:
                    shutil.rmtree(carpeta)
                    os.makedirs(carpeta)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo borrar la carpeta:\n{e}")


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
verificar_busquedas_al_inicio()
ventana.mainloop()