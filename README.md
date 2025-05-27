# PROYECTO-PARA-ASEA-1 
Este script en Python automatiza la descarga de resultados de seguimiento de guías de Correos de México (SEPOMEX) en formato PDF, utilizando Selenium WebDriver con Google Chrome  hecho durante mi servicio social en ASEA(Agencia de Seguridad Energia y Medio Ambiente)

## Funcionalidades

- Carga de archivo Excel con guías y periodos.
- Navegación automática en el portal de seguimiento de envíos de Correos de México.
- Descarga automática de PDFs y renombrado según el número de guía.
- Progreso visual y estado en tiempo real.
- Apertura automática de la carpeta de resultados al finalizar.
- Verificación de archivos existentes en la carpeta `Busquedas` al iniciar la aplicación.

## Tecnologías utilizadas

- Python 3.10+
- Tkinter (interfaz gráfica)
- Selenium (automatización de navegador)
- pandas (lectura de Excel)
- WebDriver Manager (gestión automática del ChromeDriver)

##  Requisitos

- Google Chrome instalado
- Sistema operativo: Windows, macOS o Linux
- Conexión a internet (para instalar ChromeDriver si no está presente)

## Estructura del archivo Excel

El archivo debe contener dos columnas:
1. **Guía** – Número de guía de rastreo
2. **Periodo** – Periodo correspondiente (formato usado por Sepomex)

## Cómo usar

1. Ejecuta el archivo `.py` o la versión empaquetada `.exe`.
2. Al iniciar, si existen archivos en la carpeta `Busquedas`, se preguntará si deseas eliminarlos.
3. Selecciona tu archivo Excel.
4. Haz clic en `Iniciar Proceso`.
5. Espera a que el programa descargue y renombre los archivos.
6. La carpeta `Busquedas` se abrirá automáticamente al finalizar.

