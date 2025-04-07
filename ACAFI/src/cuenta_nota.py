import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import pdfplumber
import os
import glob


def accion_nota_debito(driver, fecha_formateada, nit_emisor, xpath_accion, pdf_routes, ruta_carpeta_log):  # ingresar clientes
    """
    Función para crear una factura de compra/gasto en la página web.

    Parámetros:
    - driver: Objeto de Selenium WebDriver.
    - fecha_formateada: Fecha de elaboración en el formato correcto.
    - nit_emisor: NIT del proveedor.

    Retorno:
    - None
    """

    # encontrar el documento pdf necesario
    ruta_pdf = pdf_routes

    # Ruta de la carpeta donde buscar los PDFs
    ruta_carpeta = (ruta_carpeta_log)

    # Campo fijo a buscar
    campo_fijo = "Factura Electrónica"

    # Número de líneas a ignorar (banner)
    lineas_a_ignorar = 2  # Ajusta este valor según el número de líneas del banner

    # Abrir el archivo PDF
    with pdfplumber.open(ruta_pdf) as pdf:
        # Iterar sobre cada página del PDF
        for numero_pagina, pagina in enumerate(pdf.pages):
            # Extraer el texto de la página
            texto_pagina = pagina.extract_text()

            # Verificar si el campo fijo está en la página
            if campo_fijo in texto_pagina:
                print(f"Encontrado en la página {numero_pagina + 1}:")

                # Dividir el texto en líneas
                lineas = texto_pagina.split('\n')

                # Ignorar las primeras líneas (banner)
                lineas = lineas[lineas_a_ignorar:]

                # Buscar el campo fijo y el valor en las líneas restantes
                for linea in lineas:
                    if campo_fijo in linea:
                        # Extraer el valor que está a la derecha del campo fijo
                        partes = linea.split(campo_fijo)
                        if len(partes) > 1:  # Verificar que haya algo después del campo fijo
                            # Tomar solo el primer valor (split por espacios y tomar el primero)
                            # Solo el primer valor
                            valor = partes[1].strip().split()[0]
                            print(f"Valor encontrado: {valor}")
                        break  # Detener la búsqueda después de encontrar el valor
                print("-" * 40)
    # Si se encontró el valor, buscar el PDF en la carpeta
    if valor:
        print(
            f"\nBuscando archivos PDF que contengan '{valor}' en la carpeta '{ruta_carpeta}'...")

        # Buscar todos los archivos PDF en la carpeta
        archivos_pdf = glob.glob(os.path.join(ruta_carpeta, '*.pdf'))

        # Variable para almacenar el archivo encontrado
        archivo_encontrado = None

        # Buscar el archivo cuyo nombre contenga el valor
        for archivo in archivos_pdf:
            nombre_archivo = os.path.basename(archivo)
            if valor in nombre_archivo:
                archivo_encontrado = archivo
                print(f"Archivo encontrado: {archivo_encontrado}")

                # Extraer el primer valor antes del guion bajo (_)
                primer_valor = nombre_archivo.split('_')[0]
                print(f"Primer valor antes del '_': {primer_valor}")
                break
            if not archivo_encontrado:
                print(
                    f"No se encontró ningún archivo PDF que contenga '{valor}' en la carpeta.")
    else:
        print("No se encontró el valor en el PDF.")
    try:
        time.sleep(2)
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
        time.sleep(5)

        # Ingresar el No. de compra / Doc. Soporte
        logging.info("Seleccionando el tipo de factura...")
        no_compra = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '(//*[@class="autocompletecontainer"]//input)[1]'))
        )
        time.sleep(3)
        no_compra.send_keys(primer_valor)
        time.sleep(5)
        no_compra.send_keys(Keys.ENTER)
        time.sleep(5)
        
        # Ingresar la forma de pago
        try:
            forma_pago = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="editingAcAccount_autocompleteInput"]'))
            )
            # 2. Hacer scroll hasta el elemento
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", forma_pago)
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

        # Esperar para ver el resultado (opcional)
        time.sleep(2)

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
