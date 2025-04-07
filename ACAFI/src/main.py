from datetime import datetime
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.common.exceptions import WebDriverException, TimeoutException
import win32com.client as win32
import math
from openpyxl import load_workbook
import pandas as pd
import shutil
import time
import logging
import json
import os
import subprocess
import glob
from pathlib import Path

# Importar la función de registrar cuenta
from registrar_cuenta import registrar_cuenta_en_web
from cuenta_nota import accion_nota_debito

# funcion configurar loggin
def configurar_logging(log_file="logs/script.log"):
    """
    Configura el sistema de logging para registrar mensajes en un archivo y en la consola.

    Parámetros:
        log_file (str): Ruta del archivo donde se guardarán los logs. Por defecto es "logs/script.log".

    Raises:
        OSError: Si no se puede crear el archivo de logs o la carpeta no existe.
    """
    try:
        # Crear la carpeta 'logs' si no existe
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)

        # Configurar el logging
        logging.basicConfig(
            level=logging.INFO,  # Nivel de logging (INFO, DEBUG, ERROR, etc.)
            # Formato de los mensajes
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),  # Guardar logs en un archivo
                logging.StreamHandler()  # Mostrar logs en la consola
            ]
        )
        logging.info("Logging configurado correctamente.")
    except OSError as e:
        logging.error(f"No se pudo configurar el logging: {e}")
        raise


# Cargar configuración y convertir rutas relativas en absolutas
def cargar_configuracion(CONFIG_PATH, CREDENCIALES_PATH, DATOS_EXTRAIDOS_PATH, 
                        EXCEL_ROUTES_PATH, CONFIG_CLIENTES,BASE_DIR):
    """
    Versión mejorada que corrige los problemas de rutas y manejo de errores.

    Parámetros:
        CONFIG_PATH (str): Ruta relativa/absoluta del archivo config.json
        CREDENCIALES_PATH (str): Ruta relativa/absoluta del archivo credenciales.json
        DATOS_EXTRAIDOS_PATH (str): Ruta relativa/absoluta del archivo datos_extraidos.json
        EXCEL_ROUTES_PATH (str): Ruta relativa/absoluta del archivo excel_routes.json
        CONFIG_CLIENTES (str): Ruta relativa/absoluta del archivo clientes.json

    Retorna:
        tuple: (config, credenciales, datos_extraidos_pdf, excel_routes, config_clientes)
    """
    try:
          
        # Diccionario para mapear rutas de archivos
        PATHS = {
            'config': BASE_DIR / CONFIG_PATH,
            'credenciales': BASE_DIR / CREDENCIALES_PATH,
            'datos_extraidos': BASE_DIR / DATOS_EXTRAIDOS_PATH,
            'excel_routes': BASE_DIR / EXCEL_ROUTES_PATH,
            'clientes': BASE_DIR / CONFIG_CLIENTES
        }
        
        # Verificar existencia de archivos
        for name, path in PATHS.items():
            if not path.exists():
                raise FileNotFoundError(f"Archivo no encontrado: {path}")
        
        # Cargar todos los archivos
        with open(PATHS['config'], 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        with open(PATHS['credenciales'], 'r', encoding='utf-8') as f:
            credenciales = json.load(f)
        
        with open(PATHS['datos_extraidos'], 'r', encoding='utf-8') as f:
            datos_extraidos_pdf = json.load(f)
        
        with open(PATHS['excel_routes'], 'r', encoding='utf-8') as f:
            excel_routes = json.load(f)
        
        with open(PATHS['clientes'], 'r', encoding='utf-8') as f:
            config_clientes = json.load(f)
        
        # Convertir rutas relativas a absolutas (solo para config['paths'])
        if 'paths' in config:
            for key, path in config['paths'].items():
                if isinstance(path, str) and not os.path.isabs(path):
                    # Normalizar la ruta (convertir / a \ en Windows)
                    normalized_path = Path(path.replace('/', os.sep))
                    config['paths'][key] = str(BASE_DIR / normalized_path)
                    logging.info(f"Ruta convertida: {config['paths'][key]}")
        
        return config, credenciales, datos_extraidos_pdf, excel_routes, config_clientes

    except json.JSONDecodeError as e:
        logging.error(f"Error en formato JSON: {str(e)}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado: {str(e)}")
        raise


def iniciar_navegador(chromedriver_path, options):
    """
    Inicia el navegador Chrome utilizando el WebDriver de Chrome.

    Parámetros:
        chromedriver_path (str): Ruta al ejecutable de ChromeDriver.
        options (Options): Opciones de configuración del navegador.

    Retorna:
        WebDriver: Instancia del navegador Chrome.

    Raises:
        FileNotFoundError: Si el archivo de ChromeDriver no existe.
        WebDriverException: Si ocurre un error al iniciar el WebDriver.
        Exception: Si ocurre un error inesperado.
    """
    try:
        # Verificar si la ruta de ChromeDriver existe
        if not os.path.isfile(chromedriver_path):
            raise FileNotFoundError(
                f"ChromeDriver no encontrado en: {chromedriver_path}")

        logging.info("Iniciando el WebDriver de Chrome...")
        service = Service(executable_path=chromedriver_path)

        # Iniciar el navegador con las opciones configuradas
        driver = webdriver.Chrome(service=service, options=options)
        logging.info("WebDriver iniciado exitosamente.")
        return driver

    except FileNotFoundError as e:
        logging.error(f"Error: {e}")
        raise
    except WebDriverException as e:
        logging.error(f"Error al iniciar el WebDriver: {e}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado al iniciar el navegador: {e}")
        raise


def navegar_a_url(driver, url):
    """
    Navega a la URL especificada utilizando el WebDriver.

    Parámetros:
        driver (WebDriver): Instancia del navegador Chrome.
        url (str): URL a la que se desea navegar.

    Raises:
        WebDriverException: Si ocurre un error al navegar a la URL.
        TimeoutException: Si la página tarda demasiado en cargar.
        Exception: Si ocurre un error inesperado.
    """
    try:
        # Verificar si la URL es válida
        if not url.startswith(("http://", "https://")):
            raise ValueError(f"URL no válida: {url}")

        logging.info(f"Navegando a la URL: {url}")
        driver.get(url)  # Navegar a la URL

        # Esperar a que la página se cargue completamente
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script(
                "return document.readyState") == "complete"
        )
        logging.info("Página cargada exitosamente.")

    except ValueError as e:
        logging.error(f"Error en la URL: {e}")
        raise
    except TimeoutException as e:
        logging.error(f"La página tardó demasiado en cargar: {e}")
        raise
    except WebDriverException as e:
        logging.error(f"Error al navegar a la URL {url}: {e}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado al navegar a la URL {url}: {e}")
        raise


def login(driver, user, pas):
    """
    Realiza el proceso de login en la aplicación web.

    Parámetros:
        driver (WebDriver): Instancia del navegador Chrome.
        user (str): Nombre de usuario.
        pas (str): Contraseña.

    Raises:
        NoSuchElementException: Si no se encuentra un elemento necesario.
        TimeoutException: Si un elemento tarda demasiado en estar disponible.
        Exception: Si ocurre un error inesperado.
    """
    try:
        logging.info("Localizando el campo de usuario...")
        username_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))
        )
        shadow_user = username_element.shadow_root
        WebDriverWait(shadow_user, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#username-input"))
        ).send_keys(user)
        logging.info("Usuario ingresado correctamente.")

        logging.info("Localizando el campo de contraseña...")
        password_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '#current-password'))
        )
        shadow_pass = password_element.shadow_root
        WebDriverWait(shadow_pass, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#password-input"))
        ).send_keys(pas)
        logging.info("Contraseña ingresada correctamente.")

        logging.info("Localizando y haciendo clic en el botón de login...")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="login-submit"]'))
        ).click()
        logging.info("Login realizado exitosamente.")
    except TimeoutException as e:
        logging.error(f"Tiempo de espera agotado durante el login: {e}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado durante el login: {e}")
        raise


def cargar_excel(ruta_excel):
    """
    Carga un archivo Excel en un DataFrame de Pandas.

    Parámetros:
        ruta_excel (str): Ruta del archivo Excel a cargar.

    Retorna:
        DataFrame: DataFrame con los datos del archivo Excel.

    Raises:
        FileNotFoundError: Si el archivo no existe.
        pd.errors.EmptyDataError: Si el archivo está vacío.
        Exception: Si ocurre un error inesperado.
    """
    try:
        # Verificar si el archivo existe
        if not os.path.isfile(ruta_excel):
            raise FileNotFoundError(f"El archivo no existe: {ruta_excel}")

        logging.info(f"Cargando archivo Excel desde {ruta_excel}...")
        df = pd.read_excel(ruta_excel)

        # Verificar si el DataFrame está vacío
        if df.empty:
            raise pd.errors.EmptyDataError("El archivo Excel está vacío.")

        logging.info("Archivo Excel cargado correctamente.")
        return df

    except FileNotFoundError as e:
        logging.error(f"Error: {e}")
        raise
    except pd.errors.EmptyDataError as e:
        logging.error(f"Error: {e}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado al cargar el archivo Excel: {e}")
        raise

# converit a str los valores que figuran com oenteros


def convertir_a_str(valor):
    """
    Convierte un valor en una cadena de texto (str), eliminando ".0" si es un float sin decimales.

    Parámetros:
        valor (int, float, str, etc.): Valor a convertir.

    Retorna:
        str: Representación en cadena del valor, sin ".0" si es un float sin decimales.

    Ejemplos:
        >>> convertir_a_str(42.0)
        "42"
        >>> convertir_a_str(3.14)
        "3.14"
        >>> convertir_a_str("Hola")
        "Hola"
    """
    try:
        # Si el valor es un float sin decimales, convertirlo a int y luego a str
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        # En cualquier otro caso, convertir directamente a str
        return str(valor)
    except Exception as e:
        logging.error(f"Error al convertir el valor a str: {e}")
        raise

# procesar las variables de excel


def procesar_fila_excel(row):
    """
    Función para procesar una fila del archivo Excel y extraer los datos necesarios.

    Parámetros:
    - row: Una fila del DataFrame (serie de pandas).

    Retorno:
    - datos_fila: Un diccionario con los datos procesados de la fila.
    """
    try:
        # 🛠️ Manejo de CUFE/CUDE
        cufe = convertir_a_str(row["CUFE/CUDE"])

        # 🛠️ Manejo de Prefijo y Consecutivo
        consecutivo = row["Folio"]
        prefijo = row["Prefijo"]
        if pd.isna(prefijo) or str(prefijo).lower() == "nan":
            prefijo="FE"
        if pd.isna(consecutivo) or str(consecutivo).lower() == "nan":
            consecutivo = ""
        elif pd.isna(prefijo) or prefijo == "":
            valor = convertir_a_str(consecutivo)
            match = re.match(r"([A-Za-z]*)(\d*)", valor)
            prefijo = match.group(1)
            consecutivo = match.group(2)

        # Crear número de factura
        factura = convertir_a_str(prefijo) + convertir_a_str(consecutivo)

        # Extraer otros datos de la fila
        centro_costo_excel = (row["centro de costos"])
        fecha = convertir_a_str(row["Fecha Emisión"])
        iva = convertir_a_str(row["IVA"])
        codigo_producto = convertir_a_str(row["codigo de producto"])
        tipo_documento = convertir_a_str(row["Tipo de documento"])
        grupo = convertir_a_str(row["Grupo"])
        valor_total = convertir_a_str(row["Total"])
        # Datos del vendedor
        nit_emisor = convertir_a_str(row["NIT Emisor"])
        razon_social_vendedor = convertir_a_str(row["Nombre Emisor"])

        # Datos del receptor
        nombre_receptor = convertir_a_str(row["Nombre Receptor"])
        nit_receptor = convertir_a_str(row["NIT Receptor"])
        
        if grupo == "Emitido":
            nit_tercero = nit_receptor
        elif grupo == "Recibido":
            nit_tercero = nit_emisor
        if prefijo == "nan":
            # Separar letras y números en "prefijo"
            prefijo = re.sub(
                '[^A-Za-z]', '', prefijo) if pd.notna(prefijo) else pd.NA
            consecutivo = re.sub(
                '[^0-9]', '', prefijo) if pd.notna(prefijo) else pd.NA

            prefijo = pd.NA

        # Retornar los datos procesados
        return  cufe, factura, fecha, iva, codigo_producto, nit_tercero, razon_social_vendedor, nombre_receptor, prefijo, consecutivo, tipo_documento, centro_costo_excel, valor_total

    except Exception as e:
        logging.error(f"Error al procesar la fila: {e}")
        return None

# Función para obtener la información de un NIT


def obtener_informacion_por_nit(nit, config_clientes, centro_costo_excel):
    """
    Obtiene la información de un cliente a partir de su NIT.

    Parámetros:
        nit (str): NIT del cliente.
        config_clientes (dict): Diccionario con la configuración de los clientes.
        centro_costo_excel (str): Centro de costo obtenido del archivo Excel.

    Retorno:
        tuple: Una tupla con el nombre, centro de costo y IVA del cliente.
               Si el NIT no existe, retorna (None, None, None).

    """
    try:
        # Verificar si el NIT existe en el JSON
        if nit in config_clientes:
            # Acceder a la información del NIT
            informacion = config_clientes[nit]
            nombre = informacion["nombre"]
            centro_costo = informacion["centro de costo"]
            iva_cliente = informacion["iva"]
            codigo_iva = informacion["codigo_iva"]
            # Manejar el centro de costo
            if centro_costo == "nulo":
                centro_costo = ""
            elif centro_costo == "varios":
                centro_costo = centro_costo_excel

            logging.info(
                f"Información encontrada para el NIT {nit}: {nombre}, {centro_costo}, {iva_cliente}")
            return nombre, centro_costo, iva_cliente,codigo_iva
        else:
            logging.warning(
                f"El NIT {nit} no existe en la configuración de clientes.")
            return None, None, None, None

    except KeyError as e:
        logging.error(
            f"Error en la estructura de la configuración de clientes: {e}")
        return None, None, None
    except Exception as e:
        logging.error(f"Error inesperado al obtener información por NIT: {e}")
        return None, None, None
# condicin para notas o facturas


def contiene_nota(texto):
    """
    Función que verifica si la palabra "nota" (con o sin tildes y en cualquier caso)
    está presente en el texto proporcionado y devuelve el XPath correspondiente.

    Parámetros:
    texto (str): El texto en el que se buscará la palabra "nota".

    Retorna:
    bool: True si la palabra "nota" está presente, False en caso contrario.
    str: El XPath correspondiente a la acción a crear.
    str: Un mensaje descriptivo sobre el resultado de la búsqueda.
    """
    try:
        # Expresión regular para buscar la palabra "nota" con o sin tildes y en cualquier caso
        patron = re.compile(r'\bnota\b', re.IGNORECASE | re.UNICODE)

        # Definimos los XPaths según el tipo de acción
        xpath_nota_debito = "a[data-value='Nota débito (compras)']"
        xpath_factura_compra = "a[data-value='Factura de compra / Gasto']"

        # Buscar la palabra en el texto
        if patron.search(texto):
            # Si se encuentra la palabra "nota", devolvemos True, el XPath de "Nota débito" y un mensaje
            logging.info("La palabra 'nota' fue encontrada en el texto.")
            return True, xpath_nota_debito, "El texto contiene la palabra 'nota'. Se usará el XPath para Nota débito."
        else:
            # Si no se encuentra la palabra "nota", devolvemos False, el XPath de "Factura de compra" y un mensaje
            logging.info("La palabra 'nota' NO fue encontrada en el texto.")
            return False, xpath_factura_compra, "El texto NO contiene la palabra 'nota'. Se usará el XPath para Factura de compra."

    except Exception as e:
        # Capturamos cualquier excepción que ocurra durante la ejecución
        logging.error(f"Ocurrió un error al procesar el texto: {e}")
        return False, None, f"Error al procesar el texto: {e}"

# ejecutar scrip de pdfs


def ejecutar_script_pdf(script_pdf_path, pdf_routes_json_path, pdf_routes):
    """
    Ejecuta un script de manejo de PDFs y guarda la ruta del PDF en un archivo JSON.

    Parámetros:
        script_pdf_path (str): Ruta del script de manejo de PDFs.
        pdf_routes_json_path (str): Ruta del archivo JSON donde se guardará la ruta del PDF.
        pdf_routes (str): Ruta del archivo PDF a procesar.

    Retorno:
        str: Salida estándar del script de PDF.

    Raises:
        FileNotFoundError: Si el script de PDF o el archivo JSON no existen.
        json.JSONDecodeError: Si el archivo JSON está mal formado.
        subprocess.CalledProcessError: Si el script de PDF falla.
        Exception: Si ocurre un error inesperado.
    """
    try:
        # Verificar si el script de PDF existe
        if not os.path.isfile(script_pdf_path):
            raise FileNotFoundError(
                f"El script de PDF no existe: {script_pdf_path}")

        logging.info("Guardando la ruta del PDF en el archivo JSON...")
        # 📄 Cargar o crear el JSON
        if os.path.exists(pdf_routes_json_path):
            with open(pdf_routes_json_path, "r", encoding="utf-8") as f:
                data = json.load(f)  # Cargar contenido existente
        else:
            data = {}  # Crear estructura vacía si no existe

        # 📝 Sobrescribir con la nueva ruta
        data["path_pdf"] = pdf_routes

        # Guardar cambios en el JSON
        with open(pdf_routes_json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        logging.info(f"Ruta del PDF guardada en {pdf_routes_json_path}.")

        # Ejecutar el script de PDF
        logging.info("Ejecutando script de manejo de PDFs...")
        result = subprocess.run(
            ["python", script_pdf_path], check=True, text=True, capture_output=True
        )
        logging.info("Script de PDF ejecutado exitosamente.")
        return result.stdout

    except FileNotFoundError as e:
        logging.error(f"Error: {e}")
        raise
    except json.JSONDecodeError as e:
        logging.error(f"Error al decodificar el archivo JSON: {e}")
        raise
    except subprocess.CalledProcessError as e:
        logging.error(f"Error al ejecutar el script de PDF: {e}")
        raise
    except Exception as e:
        logging.error(f"Error inesperado al ejecutar el script de PDF: {e}")
        raise


def ingresar_cliente(driver, nit_cliente , ingreso_realizado):  # ingresar clientes
    """
    Función para ingresar un cliente en una tabla de una interfaz web.

    Parámetros:
    - driver: Objeto de Selenium WebDriver para interactuar con el navegador.
    - nit_cliente : NIT del cliente que se desea buscar y seleccionar.
    - ingreso_realizado: Bandera booleana que indica si el ingreso ya se realizó previamente.

    Retorno:
    - True: Si el ingreso se realizó correctamente.
    - False: Si el ingreso ya se había realizado o si no se encontró el cliente.
    """

    # Verificar si el ingreso ya se realizó previamente
    if not ingreso_realizado:
        try:
            # Registrar en logs que se está buscando el campo de clientes
            logging.info("Localizando el campo de mis clientes...")

            # Esperar a que el elemento que contiene la tabla de clientes esté presente en la página
            clientes_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@style="z-index: 1;"]'))
            )

            # Acceder al Shadow DOM del elemento que contiene la tabla
            shadow_clientes = clientes_element.shadow_root

            # Buscar la tabla dentro del Shadow DOM
            tabla = shadow_clientes.find_element(
                By.CSS_SELECTOR, '#wc-data-table-general')
            # Esperar 2 segundos para asegurar que la tabla esté cargada
            time.sleep(2)

            # Encontrar todas las filas de la tabla
            filas = tabla.find_elements(By.CSS_SELECTOR, 'tr')
            encontrado = False  # Bandera para indicar si se encontró el cliente
            # Convertir el NIT a cadena para comparar
            nit_adquiriente = str(nit_cliente )

            # Recorrer cada fila de la tabla
            for fila in filas:

                # Esperar 1 segundo para evitar problemas de rendimiento
                time.sleep(1)
                # Verificar si el NIT del cliente está en la fila actual
                if nit_adquiriente in fila.text:
                    try:
                        # Buscar el botón "Ingresar" dentro de la fila
                        boton_ingresar = fila.find_element(
                            By.CSS_SELECTOR, 'siigo-button-dropdown-atom')

                        # Acceder al Shadow DOM del botón
                        shadow_boton = boton_ingresar.shadow_root

                        # Buscar el botón dentro del Shadow DOM y hacer clic en él
                        boton = shadow_boton.find_element(
                            By.CSS_SELECTOR, '.button-dropdown__btn')
                        boton.click()
                        logging.info(
                            "Botón 'Ingresar' clickeado exitosamente.")

                        # Marcar como encontrado y salir del bucle
                        encontrado = True
                        break
                    except Exception as e:
                        # Registrar un error si no se puede hacer clic en el botón
                        logging.error(
                            f"Error al hacer clic en el botón 'Ingresar': {e}")
                else:
                    # Registrar que el NIT no se encontró en la fila actual
                    logging.info(
                        f"No se encontró: {nit_adquiriente} en la tabla")

            # Si no se encontró el cliente, registrar el hecho
            if not encontrado:
                logging.info("Valor no encontrado en la tabla.")

            # Registrar que el cliente se ingresó correctamente
            logging.info("Cliente ingresado correctamente.")
            time.sleep(1)  # Esperar 1 segundo antes de continuar

            # Cambiar la bandera para evitar que se repita el proceso
            ingreso_realizado = True
            return True  # Retornar True para indicar éxito

        except Exception as e:
            # Registrar un error si ocurre una excepción durante el proceso
            # Cerrar el navegador en caso de error
            logging.error(f"Error durante el ingreso de cliente: {e}")
            raise  # Relanzar la excepción para manejo externo

    # Retornar False si el ingreso ya se había realizado o si no se encontró el cliente
    return False


# Formatear fechas
def formatear_fecha(fecha, formatos=["%d-%m-%Y", "%d/%m/%Y"]):
    # Si la fecha es un Timestamp, convertirla directamente
    if isinstance(fecha, pd.Timestamp):
        return fecha.strftime("%d/%m/%Y")

    # Si la fecha ya es un string, intentar parsearla con los formatos dados
    if isinstance(fecha, str):
        for formato in formatos:
            try:
                fecha_obj = datetime.strptime(fecha, formato)
                return fecha_obj.strftime("%d/%m/%Y")
            except ValueError:
                continue

    logging.error(f"Formato de fecha no válido: {fecha}")
    return None


# ingresar datos para crear la factura de compra
def crear_factura_compra(driver, fecha_formateada, nit_tercero, xpath_accion):
    """
    Función para crear una factura de compra/gasto en la página web.

    Parámetros:
    - driver: Objeto de Selenium WebDriver.
    - fecha_formateada: Fecha de elaboración en el formato correcto.
    - nit_tercero: NIT del proveedor.

    Retorno:
    - None
    """
    try:
        # ------------------------ Interacción dentro de la página ------------------------

        # Click en crear
        logging.info("Intentando hacer clic en el botón 'Crear'...")
        time.sleep(2)
        banner_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "siigo-header-molecule.data-siigo-five9"))
        )
        shadow_banner = banner_element.shadow_root
        crear_element = shadow_banner.find_element(
            By.CSS_SELECTOR, "siigo-button-atom[data-id='header-create-button']")
        shadow_crear = crear_element.shadow_root
        shadow_crear.find_element(
            By.CSS_SELECTOR, "button[type='button'].btn-element").click()
        logging.info("Se ha dado clic en el botón 'Crear' correctamente.")
        time.sleep(1)

        # Click en factura de compra / Gasto
        logging.info("Intentando hacer clic en 'Factura de compra / Gasto'...")
        shadow_banner.find_element(
            By.CSS_SELECTOR, xpath_accion).click()
        logging.info(
            "Clic en 'Factura de compra / Gasto' realizado correctamente.")
        time.sleep(1)

        # Ingresar el TIPO DE VALOR
        logging.info("Seleccionando el tipo de factura...")
        select_tipo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[@value='ERPDocumentTypeID']/select"))
        )
        dropdown = Select(select_tipo)
        time.sleep(5)
        dropdown.select_by_visible_text("FC - 1 - Compra")
        logging.info("Se ha seleccionado la opción 'FC - 1 - Compra' en tipo.")
        time.sleep(1)

        # Ingresar la fecha de elaboración
        logging.info("Ingresando la fecha de elaboración...")
        campo_fecha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '(//*[@class="dx-texteditor-input-container"]/input)[1]'))
        )
        campo_fecha.click()
        time.sleep(0.5)
        campo_fecha.clear()
        time.sleep(1)
        campo_fecha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '(//*[@class="dx-texteditor-input-container"]/input)[1]'))
        )
        time.sleep(1)
        campo_fecha.click()
        time.sleep(2)
        driver.execute_script(
            "arguments[0].value = arguments[1];", campo_fecha, fecha_formateada)
        time.sleep(0.5)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('input'));", campo_fecha)
        time.sleep(0.5)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change'));", campo_fecha)
        time.sleep(0.5)
        logging.info("Fecha de elaboración ingresada correctamente.")

        # Ingresar el proveedor
        logging.info("Ingresando el NIT del proveedor...")
        action = ActionChains(driver)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '(//*[@class="autocompletecontainer"]/div/input)[1]'))
        ).send_keys(nit_tercero)
        time.sleep(5)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(2)
        logging.info("Proveedor ingresado correctamente.")

    except Exception as e:
        logging.error(
            f"Error durante la creación de la factura de compra: {e}")
        raise


def ingresar_datos_factura(driver, prefijo, consecutivo, codigo_producto, nit_tercero, valor, iva,iva_cliente, centro_costo,valor_total,codigo_iva):  # datos facturas
    """
    Función para ingresar los datos de una factura en la página web.
    Parámetros:
    - driver: Objeto de Selenium WebDriver.
    - prefijo: Prefijo del número de factura.
    - consecutivo: Consecutivo del número de factura.
    - lista_centros_costos: Lista de centros de costos.
    - codigo_producto: Código del producto.
    - nit_tercero: NIT del emisor.
    - valor: Valor unitario del producto.

    Retorno:
    - None
    """
    try:
        # Ingresar prefijo de número de factura proveedor
        try:
            campo_prefijo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="txtExternalPrefix"]'))
            )
            campo_prefijo.click()
            time.sleep(0.5)
            campo_prefijo.clear()
            campo_prefijo.send_keys(prefijo)
            time.sleep(1)
            logging.info(
                "Prefijo de número de factura ingresado correctamente.")
        except Exception as e:
            logging.error(
                f"Error al ingresar el prefijo de número de factura: {e}")
            raise

        # Ingresar consecutivo de número de factura proveedor
        try:
            campo_consecutivo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="txtExternalConsecutive"]'))
            )
            campo_consecutivo.send_keys(consecutivo)
            time.sleep(1)
            logging.info(
                "Consecutivo de número de factura ingresado correctamente.")
        except Exception as e:
            logging.error(
                f"Error al ingresar el consecutivo de número de factura: {e}")
            raise

        # Ingresar centro de costos
        try:
            campo_centro_costos = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '(//*[@class="autocompletecontainer"]/div/input)[2]'))
            )
            texto_a_ingresar = str(centro_costo)
            campo_centro_costos.send_keys(texto_a_ingresar)
            time.sleep(5)
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(1)

            # Verificar que el texto se haya ingresado correctamente
            texto_ingresado = campo_centro_costos.get_attribute("value")
            if texto_ingresado == texto_a_ingresar:
                logging.info("Centro de costos ingresado correctamente.")
            else:
                logging.error(
                    f"Error: El texto ingresado no coincide. Esperado: {texto_a_ingresar}, Obtenido: {texto_ingresado}")
                raise ValueError(
                    "El texto ingresado no coincide con el esperado.")

        except Exception as e:
            logging.error(f"Error al ingresar el centro de costos: {e}")
            raise

        # Ingresar tipo de producto
        try:
            select_activo_fijo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="trEditRow"]/div[2]/siigo-dropdownenum/select'))
            )
            dropdown = Select(select_activo_fijo)
            time.sleep(2)
            dropdown.select_by_visible_text("Gasto / Cuenta contable")
            logging.info("Activo fijo seleccionado correctamente.")
            time.sleep(1)
        except Exception as e:
            logging.error(f"Error al seleccionar el activo fijo: {e}")
            raise

        # Ingresar el producto a buscar
        try:
            if codigo_producto == "nan":
                logging.error(
                    f"No se encuentra un código de producto para el NIT {nit_tercero}")
                raise ValueError("Código de producto no válido.")

            select_producto = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '(//*[@class="autocompletecontainer"]/div/input)[4]'))
            )
            time.sleep(2)
            select_producto.send_keys(str(codigo_producto))
            logging.info("Producto seleccionado correctamente.")
            time.sleep(1)
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(1)
        except Exception as e:
            logging.error(f"Error al seleccionar el producto: {e}")
            raise

        # Ingresar el valor unitario
        try:
            valor_unitario = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '(//*[@class="dx-texteditor-container"]/div/input[@id="inputDecimal_siigoInputDecimal"])[3]'))
            )
            valor_unitario.click()
            time.sleep(1)
            valor_unitario.clear()
            time.sleep(1)
            # Simular escritura humana
            for char in str(valor):
                valor_unitario.send_keys(char)
                time.sleep(0.1)
            time.sleep(1)
            logging.info("Valor ingresado correctamente.")
        except Exception as e:
            logging.error(f"Error al ingresar el valor unitario: {e}")
            raise

        # sacar el iva correspondiente y ver si esta vacio o no
        try:

            # Verificar si la variable es NaN o un string vacío/contiene solo espacios/guiones
            if not (math.isnan(iva) if isinstance(iva, (int, float)) else not iva.strip().replace('-', '').strip()):
                if iva != '0' and iva != 0:
                    # Localizar el elemento <select>
                    diferencia_iva = round(float(valor) * 0.19, 2)
                    resultado_iva = float(valor) + float(diferencia_iva)
                    print(float(valor_total))
                    if resultado_iva != float(valor_total):
                        WebDriverWait(driver,10).until(EC.visibility_of_element_located(
                            (By.XPATH,'//*[text()=" Agregar otro ítem "]')
                        )).click()
                        # Ingresar tipo de producto
                        try:
                            select_activo_fijo = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, '//*[@id="trEditRow"]/div[2]/siigo-dropdownenum/select'))
                            )
                            dropdown = Select(select_activo_fijo)
                            time.sleep(2)
                            dropdown.select_by_visible_text("Gasto / Cuenta contable")
                            logging.info("Activo fijo seleccionado correctamente.")
                            time.sleep(1)
                        except Exception as e:
                            logging.error(f"Error al seleccionar el activo fijo: {e}")
                            raise
                        # Ingresar el producto a buscar
                        try:
                            if codigo_iva == "nan":
                                logging.error(
                                    f"No se encuentra un código de producto para el NIT {codigo_iva}")
                                raise ValueError("Código de producto no válido.")

                            select_producto = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, '(//*[@class="autocompletecontainer"]/div/input)[4]'))
                            )
                            time.sleep(2)
                            select_producto.send_keys(str(codigo_iva))
                            logging.info("Producto seleccionado correctamente.")
                            time.sleep(1)
                            ActionChains(driver).send_keys(Keys.ENTER).perform()
                            time.sleep(1)
                        except Exception as e:
                            raise
                    
                        # Ingresar el valor unitario
                        try:
                            valor_unitario = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable(
                                    (By.XPATH, '(//*[@class="dx-texteditor-container"]/div/input[@id="inputDecimal_siigoInputDecimal"])[3]'))
                            )
                            valor_unitario.click()
                            time.sleep(1)
                            valor_unitario.clear()
                            time.sleep(1)
                            # Simular escritura humana
                            for char in str(iva):
                                valor_unitario.send_keys(char)
                                time.sleep(0.1)
                            time.sleep(1)
                            
                            logging.info("Valor ingresado correctamente.")
                        except Exception as e:
                            logging.error(f"Error al ingresar el valor unitario: {e}")
                            raise
                    else:
                        # Cambia "ID_DEL_SELECT" por el ID real del elemento
                        select_element = driver.find_element(By.XPATH,
                            '//siigo-dropdown[@id="editAddTax"]//*[@id="dropdown_dropdownSelect"]')
                        # Obtener todas las opciones del <select>
                        options_select_iva = select_element.find_elements(By.TAG_NAME, "option")
                        # Recorrer las opciones y encontrar el texto "IVA 19 MV"
                        for option in options_select_iva:
                            print(option)
                            if iva_cliente in option.text:
                                logging.info(f"Texto encontrado: {option.text}")
                                # Aquí puedes realizar la acción que necesites, como seleccionar la opción
                                option.click()  # Selecciona la opción
                                break
                    
            else:
                logging.warning(
                    "La variable es NaN o está vacía/contiene solo espacios/guiones, no se realiza ninguna acción.")

        except Exception as e:
            # Capturar cualquier excepción y registrar el error
            logging.error(f"Ocurrió un error: {e}", exc_info=True)
            raise
        # Ingresar la forma de pago
        try:
            forma_pago = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="editingAcAccount_autocompleteInput"]'))
            )
            
            # 2. Hacer scroll hasta el elemento
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", forma_pago)
            time.sleep(1)
            
            # 3. Esperar un breve momento (opcional, pero recomendado para páginas lentas)
            time.sleep(1)

            # 4. Hacer clic (usando JavaScript para evitar intercepciones)
            forma_pago.click()

            elemento = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//table[contains(@class, 'siigo-ac-table')]//div[text()=' Otras cuentas por pagar ']"))
            )
            # Desplazarse hasta el elemento (si es necesario)
            ActionChains(driver).move_to_element(elemento).perform()
            # Hacer clic en el elemento
            elemento.click()
            logging.info("Forma de pago seleccionada correctamente.")
        except Exception as e:
            logging.error(f"Error al seleccionar la forma de pago: {e}")
            raise

        try:
            # Click en la X si aparece
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                By.XPATH, '//*[@class="icon-siigo-simbolos-cerrar red"]'))).click()
            logging.info("Ventana emergente cerrada correctamente.")
        except Exception as e:
            logging.error(f"Error al cerrar la ventana emergente:")

        time.sleep(2)

        try:
            # Hacer click en guardar
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[contains(@class,"SiigoButtonPrimary")]'))).click()
            logging.info("Factura guardada correctamente.")
            time.sleep(2)
        except Exception as e:
            logging.error(f"Error al guardar la factura: {e}")
            raise

        # Esperar y obtener el texto de la factura
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@class="title-container"]'))
        )
        logging.info("Factura procesada correctamente.")

    except Exception as e:
        logging.error(
            f"Error general al ingresar los datos de la factura: {e}")
        raise


# funcion para obtener los datos de la factura
def obtener_y_mover_factura(driver, output_folder, pdf_routes, razon_social_vendedor, factura, ruta_carpeta_log):
    """
    Obtiene el número de factura de una página web y mueve el archivo PDF correspondiente.

    :return: (numero_factura, True) si la operación fue exitosa, (None, False) en caso contrario.
    """
    TIMEOUT = 30  # Tiempo máximo de espera para el texto de la factura

    try:
        logging.info("Esperando el texto de la factura...")
        # Esperar a que aparezca el texto de la factura
        texto_factura_compra = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@class="title-container"]'))
        )
        time.sleep(5)
        texto_factura_compra = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@class="title-container"]'))
        ).text

        # Expresión regular para extraer el número de factura
        match = re.search(r':\s*(\S+)', texto_factura_compra)
        if not match:
            logging.error("No se pudo extraer el número de factura.")
            return None, False

        numero_factura = match.group(1)
        logging.info(f"Número de factura extraído: {numero_factura}")

        # Asegurar que la carpeta destino existe
        os.makedirs(ruta_carpeta_log, exist_ok=True)

        # Verificar si el PDF original existe
        if not os.path.exists(pdf_routes):
            logging.warning(f"No se encontró el PDF: {pdf_routes}")
            return None, False,None

        # Nueva ruta con el número de factura como nombre
        new_pdf_path = os.path.join(
            ruta_carpeta_log, f"{numero_factura}_{razon_social_vendedor}_{factura}.pdf")
        output_pdf_path = os.path.join(
            ruta_carpeta_log, f"{numero_factura}")
        # Verificar si el archivo ya existe en la nueva ubicación
        if os.path.exists(new_pdf_path):
            logging.warning(
                f"El archivo {new_pdf_path} ya existe. No se moverá.")
            return None, False,None

        # Mover y renombrar el archivo
        shutil.move(pdf_routes, new_pdf_path)
        logging.info(f"PDF movido a: {new_pdf_path}")

        return numero_factura, True,ruta_carpeta_log

    except Exception as e:
        logging.error(f"Error inesperado al obtener el número de factura: {e}")

    return None, False,None # En caso de error, retorna siempre una tupla


def enviar_correos(nombre_archivo, lista_correos):
    try:
        # 📩 Crear una instancia de Outlook
        outlook = win32.Dispatch('Outlook.Application')

        # Verificar si el archivo a adjuntar existe
        ruta_archivo = fr"C:\Users\santi\OneDrive\Escritorio\Swith_bots\Swith_bots\ACAFI\inputs\{nombre_archivo}"
        if not os.path.exists(ruta_archivo):
            print(f"⚠ Archivo no encontrado: {ruta_archivo}")
            return

        # Iterar sobre cada dirección de correo en la lista
        for correo in lista_correos:
            try:
                mail = outlook.CreateItem(0)  # 0 representa un correo nuevo

                # ✉ Configurar el correo
                mail.Subject = "Notificación: Ejecución del Bot Finalizada"
                mail.Body = "El bot ha finalizado su ejecución. Todas las filas han sido procesadas o se alcanzó el límite de 3 intentos."
                mail.To = correo

                # Adjuntar el archivo
                mail.Attachments.Add(ruta_archivo)
                print(f"Archivo adjuntado correctamente para {correo}.")

                # 📤 Enviar el correo
                mail.Send()
                print(f"✅ Correo electrónico enviado correctamente a {correo}.")
            except Exception as e:
                print(f"❌ Error al enviar el correo electrónico a {correo}: {e}")
    except Exception as e:
        print(f"❌ Error general al enviar los correos electrónicos: {e}")


# ----------------------------
# EJECUCIÓN PRINCIPAL DEL SCRIPT
# ----------------------------
if __name__ == "__main__":
    
    ###########################################################
    # Configurar el sistema de logging para registrar mensajes
    ###########################################################
    configurar_logging()
    logging.info("Iniciando la ejecución del script principal.")

    # Contador de ejecuciones
    ejecuciones_maximas = 3
    ejecuciones_realizadas = 0
    
    ###########################################################
    # Definir las rutas de los archivos de configuración y credenciales
    ###########################################################
    CONFIG_PATH = "config/config.json"
    CREDENCIALES_PATH = "config/credenciales.json"
    DATOS_EXTRAIDOS_PATH = "config/datos_extraidos.json"
    EXCEL_ROUTES_PATH = "config/ruta_excel.json"
    CONFIG_CLIENTES = "config/configuracion_usuarios.json"
    ###########################################################
    # Cargar la configuración, credenciales y datos extraídos
    ###########################################################
    
    # Obtener ruta base del proyecto (ACAFI/)
    BASE_DIR = Path(__file__).parent.parent
    
    config, credenciales, datos_extraidos_pdf, excel_routes, config_clientes = cargar_configuracion(
        CONFIG_PATH, CREDENCIALES_PATH, DATOS_EXTRAIDOS_PATH, EXCEL_ROUTES_PATH, CONFIG_CLIENTES, BASE_DIR
    )
    logging.info(
        "Configuración y credenciales cargadas correctamente.")

    
    carpeta = config["paths"]["inputs"]
    config_folder = config["paths"]["config"]
    
    try:
        for archivo in os.listdir(carpeta):
            if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
                ruta_archivo = os.path.join(carpeta, archivo)
                try:
                    df = pd.read_excel(ruta_archivo, engine="openpyxl")
                    nombre_archivo = ruta_archivo.split("\\")[-1]
                    # Extrae el NIT (todo antes del primer '(', '_' o '.')
                    nit_receptor = re.split(r'[_(.]', nombre_archivo)[0]
                    # Bandera para controlar si todas las filas han sido procesadas
                    todas_filas_procesadas = False
                    
                    # Definir tamaño de lote
                    TAMANO_LOTE = 5
                    # Archivo para guardar progreso
                    ARCHIVO_PROGRESO = os.path.join(config_folder, f"progreso.json")
                    
                     # Limpiar el archivo de progreso si existe
                    if os.path.exists(ARCHIVO_PROGRESO):
                        os.remove(ARCHIVO_PROGRESO)
                        logging.warning("¡Se eliminó el archivo de progreso previo!")

                    while ejecuciones_realizadas < ejecuciones_maximas and not todas_filas_procesadas:
                        ejecuciones_realizadas += 1
                        logging.info(f"Ejecución número {ejecuciones_realizadas}.")
                        try:
                            ###########################################################
                            # Configurar las opciones de Chrome para el navegador
                            ###########################################################
                            options = Options()
                            # Maximizar la ventana del navegador
                            options.add_argument("--start-maximized")
                            # Deshabilitar la política de mismo origen
                            options.add_argument("--disable-web-security")
                            # Deshabilitar notificaciones
                            options.add_argument("--disable-notifications")
                            # Evitar detección de automatización
                            options.add_argument(
                                "--disable-blink-features=AutomationControlled")
                            # Ignorar errores de certificados SSL
                            options.add_argument("--ignore-certificate-errors")
                            # Permitir conexiones inseguras a localhost
                            options.add_argument("--allow-insecure-localhost")
                            # evitar que el navegador se cierre para ver el error
                            options.add_experimental_option("detach", True)

                            logging.info("Opciones del navegador configuradas correctamente.")

                            ###########################################################
                            # Iniciar el navegador Chrome con las opciones configuradas
                            ###########################################################
                            driver = iniciar_navegador(config["paths"]["web_driver"], options)
                            logging.info("Navegador iniciado correctamente.")

                            ###########################################################
                            # Navegar a la URL de la página principal
                            ###########################################################
                            navegar_a_url(driver, config["urls"]["main"])
                            logging.info(
                                f"Navegado a la URL principal: {config['urls']['main']}")
                            
                            ###########################################################
                            # Cargar el archivo Excel que contiene los datos a procesar
                            ###########################################################
                            df = cargar_excel(ruta_archivo)
                            logging.info(
                            f"Archivo Excel cargado correctamente: {ruta_archivo}")
                            
                            # Extraer solo el nombre del archivo sin la extensión
                            nombre_archivo = excel_routes['ruta_archivo.excel'].split(
                                "\\")[-1]
                            nit_cliente  = re.split(r'[_(.]', nombre_archivo)[0]
                            
                            
                            # Verificar si la columna 'PDF Generado' existe, si no, crearla
                            if 'PDF Generado' not in df.columns:
                                df['PDF Generado'] = 'No'
                            # Definir las columnas necesarias
                            columnas_necesarias = ['PDF Generado', 'Procesamiento Exitoso',
                                                'Forma de Pago', 'Nombre PDF', 'Mensaje Error']

                            # Verificar si las columnas existen, si no, crearlas con valores vacíos
                            for col in columnas_necesarias:
                                if col not in df.columns:
                                    df[col] = ""  # Se inicializan vacías
                                    
                            # Verificar si todas las filas ya están procesadas
                            if all(df['PDF Generado'] == 'Sí'):
                                logging.info(
                                    "Todas las filas ya están procesadas. Finalizando ejecución.")
                                todas_filas_procesadas = True
                                break
                            else: 
                                logging.info(
                                    "todavia faltan documentos por generar"
                                )
                            ###########################################################
                            # Iniciar sesión en la aplicación web
                            ###########################################################
                            login(driver, credenciales[nit_cliente ]["usuario"], credenciales[nit_cliente ]["contrasena"])
                            logging.info("Sesión iniciada correctamente.")
                            
                            # Bandera para controlar si el ingreso ya se realizó
                            ingreso_realizado = False
                            
                            # Cargar o inicializar progreso
                            if os.path.exists(ARCHIVO_PROGRESO):
                                with open(ARCHIVO_PROGRESO, 'r') as f:
                                    progreso = json.load(f)
                                lote_actual = progreso['ultimo_lote'] + 1
                            else:
                                progreso = {
                                    'archivo': ruta_archivo,
                                    'ultimo_lote': -1,
                                    'filas_procesadas': 0
                                }
                                lote_actual = 0
                                
                            # Calcular lotes pendientes
                            df_pendientes = df[df['PDF Generado'] != 'Sí']
                            total_filas = len(df_pendientes)
                            total_lotes = (total_filas + TAMANO_LOTE - 1) // TAMANO_LOTE
                            
                            # Procesar lotes pendientes
                            for lote_num in range(lote_actual, total_lotes):
                                inicio = lote_num * TAMANO_LOTE
                                fin = min(inicio + TAMANO_LOTE, total_filas)
                                lote = df_pendientes.iloc[inicio:fin]
                                
                                logging.info(f"Procesando lote {lote_num + 1}/{total_lotes} (filas {inicio+1}-{fin})")

                                ###########################################################
                                # Iterar sobre cada fila del DataFrame (archivo Excel)
                                ###########################################################
                                for index, row in lote.iterrows():  # Solo procesamos 10 filas por iteración
                                    try:
                                        logging.info(
                                            f"Procesando fila {index + 1} del archivo Excel.")

                                        ###########################################################
                                        # Extraer y procesar los datos de la fila actual del Excel
                                        ###########################################################
                                        datos_fila = procesar_fila_excel(row)
                                        if not datos_fila:
                                            logging.warning(
                                                f"Fila {index + 1} no procesada correctamente. Saltando...")
                                            continue
                                        # Desempaquetar los datos de la fila
                                        cufe, factura, fecha, iva, codigo_producto, nit_tercero, razon_social_vendedor, nombre_receptor, prefijo, consecutivo, tipo_documento, centro_costo_excel,valor_total = datos_fila
                                        logging.info("Datos extraídos correctamente de la fila.")

                                        ###########################################################
                                        # Verificar si el valor de "CUFE/CUDE" está vacío
                                        ###########################################################
                                        if pd.isna(cufe):
                                            logging.warning(
                                                f"Fila {index + 1} tiene datos incompletos (CUFE/CUDE vacío). Saltando...")

                                        output_folder = config["paths"]["output"]
                                        ##########################################################
                                        # leer los parametros del cliente de el config
                                        ####################################################
                                        nombre, centro_costo, iva_cliente,codigo_iva = obtener_informacion_por_nit(
                                            nit_cliente , config_clientes, centro_costo_excel)
                                        # Verificar si se encontró la información
                                        if nombre is not None:
                                            print(f"Información para el NIT {config_clientes}:")
                                            print("Nombre:", nombre)
                                            print("Centro de costo:", centro_costo)
                                            print("IVA:", iva_cliente)
                                        else:
                                            print(
                                                f"El NIT {config_clientes} no existe en el JSON.")
                                        ###########################################################
                                        # Formatear la fecha para el formato requerido
                                        ###########################################################
                                        # Formatear la fecha antes de usarla
                                        fecha_formateada = formatear_fecha(fecha)

                                        # Verificar y registrar la fecha formateada
                                        if fecha_formateada:
                                            logging.info(f"Fecha formateada: {fecha_formateada}")
                                        else:
                                            logging.error(
                                                f"No se pudo formatear la fecha: {fecha}")
                                        # Obtener la fecha actual
                                        ahora = datetime.now()
                                        año = ahora.strftime("%Y")
                                        mes = ahora.strftime("%m")
                                        dia = ahora.strftime("%d")

                                        ruta_carpeta_log = os.path.join(
                                            output_folder, str(nit_cliente ), año, mes, dia)

                                        ##########################################################
                                        # Construir la ruta del archivo PDF asociado al CUFE/CUDE
                                        # \
                                        pdf_folder = os.path.join(
                                            config["paths"]["output"]
                                        )
                                        pdf_routes = os.path.join(
                                            config["paths"]["pdf"],nit_cliente , f"{cufe}.pdf")
                                        pdf_routes_json_path = os.path.join(
                                            config["paths"]["config"], "pdf_routes.json")

                                        if not os.path.isfile(pdf_routes):
                                            logging.warning(
                                                f"El archivo PDF no existe en la ruta: {pdf_routes}")
                                            raise FileNotFoundError(
                                                f"El archivo PDF no existe en la ruta: {pdf_routes}")

                                        logging.info(f"Procesando archivo PDF: {pdf_routes}")

                                        ###########################################################
                                        # Ejecutar el script de manejo de PDFs
                                        ###########################################################
                                        script_pdf_path = os.path.join(
                                            BASE_DIR,"src","main_pdf.py")
                                        ejecutar_script_pdf(
                                            script_pdf_path, pdf_routes_json_path, pdf_routes)
                                        logging.info(
                                            "Script de manejo de PDFs ejecutado correctamente.")

                                        ###########################################################
                                        # Cargar el archivo JSON con los datos extraídos de los PDFs
                                        ###########################################################
                                        json_datos_extraidos = os.path.join(
                                            config["paths"]["config"], "datos_extraidos.json")
                                        with open(json_datos_extraidos, "r", encoding="utf-8") as file:
                                            datos_extraidos = json.load(file)
                                        logging.info("Datos extraídos cargados correctamente.")

                                        # Verificar si datos_extraidos es una lista y tiene al menos un elemento
                                        if isinstance(datos_extraidos, list) and len(datos_extraidos) > 0:
                                            # Acceder al primer elemento de la lista (que es un diccionario)
                                            primer_elemento = datos_extraidos[0]

                                            # Obtener la forma de pago (usando un valor predeterminado si la clave no existe)
                                            forma_de_pago = primer_elemento.get(
                                                "Forma de Pago", "Desconocido")
                                            logging.info(f"Forma de pago: {forma_de_pago}")
                                            valor = primer_elemento.get(
                                                "Total Bruto Factura", "Desconocido")
                                            logging.info(f"Total Bruto Factura: {valor}")
                                        else:
                                            logging.error(
                                                "El archivo JSON no contiene una lista válida o está vacío.")
                                            
                                        ### ------------------apartado web-------------------------###

                                        ###########################################################
                                        # Ingresar los datos del cliente receptor en la aplicación web - 1
                                        ###########################################################
                                        if not ingresar_cliente(driver, nit_cliente , ingreso_realizado):
                                            logging.warning(
                                                "El ingreso ya se había realizado o hubo un error.")

                                        logging.info(
                                            "Ingreso del cliente realizado correctamente.")
                                        ingreso_realizado = True

                                        ########################################################
                                        # función para saber si es una nota o un credito
                                        ########################################################

                                        # Llamamos a la función y almacenamos los resultados en variables específicas
                                        contiene_nota_resultado, xpath_accion, mensaje_resultado = contiene_nota(
                                            tipo_documento)

                                        # Imprimimos los resultados
                                        logging.info(
                                            f"¿Contiene la palabra 'nota'? {contiene_nota_resultado}")
                                        logging.info(f"Mensaje: {mensaje_resultado}")
                                        # Tomamos decisiones basadas en el resultado booleano
                                        if contiene_nota_resultado:
                                            # Si contiene la palabra "nota", ejecutamos la función relacionada con Nota débito
                                            accion_nota_debito(
                                                driver, fecha_formateada, nit_tercero, xpath_accion, pdf_routes, ruta_carpeta_log)

                                        else:
                                            # El resto del código continúa normalmente
                                            logging.info(
                                                "El resto del código continúa su ejecución...")
                                            print("Continuando con el flujo normal del programa...")

                                            ###########################################################
                                            # Crear factura de compra en la aplicación web -2
                                            ###########################################################
                                            crear_factura_compra(
                                                driver, fecha_formateada, nit_tercero, xpath_accion,)
                                            logging.info("Factura de compra creada correctamente.")

                                            ###########################################################
                                            # Registrar la cuenta en la aplicación web con los datos extraídos -3
                                            ###########################################################
                                            registrar_cuenta_en_web(
                                                driver, datos_extraidos, nit_tercero, razon_social_vendedor)
                                            logging.info(
                                                "Cuenta registrada correctamente en la aplicación web.")

                                            ###########################################################
                                            # Ingresar datos de la factura en la aplicación web -4
                                            ###########################################################
                                            ingresar_datos_factura(
                                                driver, prefijo, consecutivo, codigo_producto, nit_tercero, valor, iva, iva_cliente, centro_costo,valor_total,codigo_iva)
                                            logging.info(
                                                "Datos de la factura ingresados correctamente.")

                                        ###########################################################
                                        # Obtener y mover la factura generada -5
                                        ###########################################################

                                        numero_factura, resultado,output_pdf_path = obtener_y_mover_factura(
                                            driver=driver,
                                            output_folder=output_folder,
                                            pdf_routes=pdf_routes,
                                            razon_social_vendedor=razon_social_vendedor,
                                            factura=factura,
                                            ruta_carpeta_log=ruta_carpeta_log,
                                        )

                                        if resultado:
                                            logging.info("La factura se procesó correctamente.")
                                            estado_procesamiento = "Exitoso"
                                            mensaje_error = ""
                                            # Actualizar la columna 'PDF Generado' a 'Sí'
                                            df.at[index, 'PDF Generado'] = 'Sí'
                                            progreso['filas_procesadas'] += 1
                                            # "Exitoso" o "Fallido"
                                            df.at[index, 'Procesamiento Exitoso'] = 'Procesamiento Exitoso'
                                            # Guardar el archivo Excel después de cada actualización
                                            df.to_excel(
                                                excel_routes["ruta_archivo.excel"], index=False)
                                            logging.info(
                                                f"Archivo Excel actualizado en: {excel_routes['ruta_archivo.excel']}")
                                        else:
                                            logging.error(" Hubo un error al procesar la factura.")
                                            estado_procesamiento = "Fallido"
                                            mensaje_error = "Error al procesar la factura"
                                            raise Exception(mensaje_error)
                                        # Variable que ya tienes
                                        df.at[index, 'Forma de Pago'] = forma_de_pago
                                        # Mensaje de error si falló
                                        df.at[index, 'Mensaje Error'] = mensaje_error
                                        # Nombre del PDF generado
                                        df.at[index, 'Nombre PDF'] = numero_factura
                                        
                                        # Guardar el archivo Excel después de cada actualización
                                        df.to_excel(
                                            excel_routes["ruta_archivo.excel"], index=False)
                                    except Exception as e:
                                            logging.error(
                                                f"Error al procesar la fila {index + 1}: {e}")
                                            # Agregar los datos de la fila procesada al DataFrame de control con estado fallido
                                            forma_de_pago = "null"
                                            # Variable que ya tienes
                                            df.at[index, 'Forma de Pago'] = forma_de_pago
                                            # Mensaje de error si falló
                                            df.at[index, 'Mensaje Error'] = str(e)
                                            # Nombre del PDF generado
                                            df.at[index, 'Nombre PDF'] = ""
                                            # procesamiento no exitoso
                                            df.at[index, 'Procesamiento Exitoso'] = "Fallido"
                                            # Guardar el archivo Excel después de cada actualización
                                            df.to_excel(
                                                excel_routes["ruta_archivo.excel"], index=False)
                                                                # Actualizar progreso después de cada lote
                                                                
                                progreso['ultimo_lote'] = lote_num
                                with open(ARCHIVO_PROGRESO, 'w') as f:
                                    json.dump(progreso, f)    
                                
                                # Guardar cambios en el Excel
                                df.to_excel(ruta_archivo, index=False)
                                logging.info(f"Progreso guardado. Lote {lote_num + 1} completado.")
                                
                                # Verificar si se completó todo
                                if len(df[df['PDF Generado'] != 'Sí']) == 0:
                                    todas_filas_procesadas = True
                                    if os.path.exists(ARCHIVO_PROGRESO):
                                        os.remove(ARCHIVO_PROGRESO)
                                    logging.info("¡Todo el archivo Excel ha sido procesado con éxito!")
                                else:
                                    ###########################################################
                                    # Esperar 5 minutos antes de la siguiente iteración
                                    ###########################################################
                                    logging.info("Esperando 5 minutos antes de la siguiente iteración...")
                                    time.sleep(300)  # 300 segundos = 5 minutos         
                                    
                        except Exception as e:
                            logging.error(f"Error durante la ejecución del lote: {str(e)}")
                            if ejecuciones_realizadas == ejecuciones_maximas:
                                logging.error("Se alcanzó el máximo de intentos. Abortando.")
                                raise
                    ###########################################################
                    # Cerrar el navegador al finalizar
                    ###########################################################
                    if 'driver' in locals():
                        driver.quit()
                        logging.info("Navegador cerrado correctamente.")
                except Exception as e:
                    logging.error(f"Error en la ejecución principal: {e}")
                        
                # Enviar correo electrónico al finalizar
                # Extraer la lista de correos electrónicos
                correos = config.get("correos", [])
                enviar_correos(nombre_archivo,correos)

                # Mover y renombrar el archivo
                shutil.move(excel_routes["ruta_archivo.excel"], f"{ruta_carpeta_log}/{nit_cliente}.xlsx")
                logging.info(f"PDF movido a: {output_pdf_path}")
                # 2. Proceso de consolidación mensual con logging
                try:
                    archivo_log = f"{ruta_carpeta_log}/{nit_cliente}.xlsx"
                    df_nuevo = pd.read_excel(archivo_log)
                    
                    # Validación de estructura
                    if 'Fecha Emisión' not in df_nuevo.columns:
                        logging.warning("El archivo no contiene columna 'Fecha'. No se puede clasificar por mes.")
                    else:
                        fecha = pd.to_datetime(df_nuevo['Fecha'].iloc[0])
                        nombre_mes = fecha.strftime("%Y-%m")
                        archivo_mes = f"facturas_{nombre_mes}.xlsx"
                        ruta_mes = os.path.join("facturas_mensuales", archivo_mes)
                        
                        # Crear directorio si no existe
                        os.makedirs(os.path.dirname(ruta_mes), exist_ok=True)
                        
                        if os.path.exists(ruta_mes):
                            df_existente = pd.read_excel(ruta_mes)
                            
                            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
                            logging.info(f"Archivo {archivo_mes} actualizado con {len(df_nuevo)} nuevos registros")
                        else:
                            df_final = df_nuevo
                            logging.info(f"Archivo mensual {archivo_mes} creado con {len(df_nuevo)} registros")
                        
                        df_final.to_excel(ruta_mes, index=False)
                        logging.info(f"Consolidación mensual completada: {ruta_mes}")

                except Exception as e:
                    logging.error(f"Error en consolidación mensual: {str(e)}", exc_info=True)
                logging.info("Ejecución finalizada.")

    except Exception as e:
        logging.error(f"Error al procesar el archivo {ruta_archivo}: {e}")