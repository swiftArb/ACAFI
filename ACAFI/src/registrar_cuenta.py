import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from nameparser import HumanName

def registrar_cuenta_en_web(driver, datos_extraidos, nit_emisor, razon_social_vendedor):
    """
    Función para registrar una cuenta en la aplicación web.

    Parámetros:
    - driver: Objeto de Selenium WebDriver.
    - datos_extraidos: Lista de diccionarios con los datos extraídos.
    - nit_emisor: NIT del emisor.
    - razon_social_vendedor: Razón social del vendedor.

    Retorno:
    - None
    """
    # Validar datos_extraidos
    if not datos_extraidos or not isinstance(datos_extraidos, list):
        logging.error("Datos extraídos no válidos o vacíos.")
        return

    for dato in datos_extraidos:
        if not isinstance(dato, dict):
            logging.error("Elemento en datos_extraidos no es un diccionario.")
            continue

        # Extraer datos del diccionario
        archivo = dato.get('Archivo')
        tipo_contribuyente = dato.get(
            'Información del vendedor', {}).get('Tipo de contribuyente')
        regimen_fiscal = dato.get(
            'Información del vendedor', {}).get('Régimen fiscal')
        descripcion_producto = dato.get('Descripción del producto')

        # Validar datos obligatorios
        if None in (archivo, tipo_contribuyente, regimen_fiscal):
            logging.error(f"Faltan datos obligatorios en el registro: {dato}")
            continue

        try:
            # Esperar a que el modal esté presente en el DOM
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@class="modal-content"]')))
            logging.info("Creando un nuevo usuario...")

            # Ingresar el tipo de contribuyente (Empresa o Persona Natural)
            try:
                shadow_tipo = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#CO-CL-MX > div > siigo-dropdownlist-web"))
                )
                shadow_tipo_contribuyente = shadow_tipo.shadow_root

                # Hacer clic en el menú desplegable para abrirlo
                dropdown = shadow_tipo_contribuyente.find_element(
                    By.CSS_SELECTOR, '.mdc-select')
                dropdown.click()
                time.sleep(1)  # Esperar a que el menú se abra

                # Seleccionar la opción correcta según el tipo de contribuyente
                opciones = shadow_tipo_contribuyente.find_elements(
                    By.CSS_SELECTOR, 'span.mdc-list-item__text')
                for opcion in opciones:
                    if tipo_contribuyente == "Persona Jurídica" and "Empresa" in opcion.text:
                        nombre_selector = "#MX_MR_EX-CO_E-1 > div > siigo-textfield-web"
                        opcion.click()
                        break
                    elif tipo_contribuyente == "Persona Natural" and "Es persona" in opcion.text:
                        nombre_selector = "#MX_FS-CO_P-1 > div > siigo-textfield-web"
                        nombre_completo = razon_social_vendedor
                        nombre = HumanName(nombre_completo)
                        print("Nombre:", nombre.first)
                        print("Apellido:", nombre.last)
                        razon_social_vendedor = nombre.first
                        opcion.click()
                        break

                logging.info(
                    "Tipo de contribuyente seleccionado correctamente.")
                time.sleep(1)
            except Exception as e:
                logging.error(
                    f"Error al seleccionar el tipo de contribuyente: {e}")
                raise

            # Campo de identificación (NIT)
            try:
                campo_identificacion = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#CO_P_E-2 > div > siigo-identification-input-web"))
                )
                shadow_identificacion = campo_identificacion.shadow_root
                shadow_identificacion.find_element(
                    By.CSS_SELECTOR, "#identification > input").send_keys(nit_emisor)
                logging.info("Identificación ingresada correctamente.")
                time.sleep(1)
            except Exception as e:
                logging.error(f"Error al ingresar la identificación: {e}")
                raise

            # Campo de razón social
            try:
                campo_razon_social = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, nombre_selector))
                )
                shadow_razon_social = campo_razon_social.shadow_root
                shadow_razon_social.find_element(
                    By.CSS_SELECTOR, ".mdc-text-field__input").send_keys(razon_social_vendedor)
                logging.info("Razón social ingresada correctamente.")
                time.sleep(1)
            except Exception as e:
                logging.error(f"Error al ingresar la razón social: {e}")
                raise


             # Campo apellido (si aplica)
            try:
                campo_apellido = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#MX_FS-CO_P2 > div > siigo-textfield-web"))
                )
                shadow_campo_apellido  = campo_apellido .shadow_root
                shadow_campo_apellido.find_element(
                    By.CSS_SELECTOR, ".mdc-text-field__input").send_keys(nombre.last)
                logging.info("apellido ingresado")
                time.sleep(1)
            except Exception as e:
                logging.error(f"no hay que ingresar apellido")
            
            # Verificar si el régimen fiscal es "R-99-PN"
            if regimen_fiscal != "R-99-PN":
                try:
                    # Localizar el checkbox por el código (O-13, O-15, etc.)
                    checkbox = driver.find_element(
                        By.XPATH, f'//*[text()="{regimen_fiscal}"]')
                    checkbox.click()
                    logging.info(
                        f"Se hizo clic en el checkbox con código: {regimen_fiscal}")
                except Exception as e:
                    logging.error(
                        f"No se encontró el checkbox con código: {regimen_fiscal}")

            # Guardar los cambios
            try:
                guardar_shadow = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "body > modal-container > div > div > div > div.modal-footer > div > siigo-button-atom:nth-child(2)"))
                )
                guardar_shadow = guardar_shadow.shadow_root
                guardar_shadow.find_element(By.CSS_SELECTOR, "button").click()
                logging.info("Cambios guardados correctamente.")
            except Exception as e:
                logging.error("Error al hacer clic en 'Guardar'.")
            time.sleep(3)
            # Ingresar el proveedor
            logging.info("Ingresando el NIT del proveedor...")
            action = ActionChains(driver)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '(//*[@class="autocompletecontainer"]/div/input)[1]'))
            ).send_keys(nit_emisor)
            time.sleep(1)
            action.send_keys(Keys.ENTER).perform()
            time.sleep(5)
            logging.info("Proveedor ingresado correctamente.")

        except Exception as e:
            logging.info(f"no es necesario crear un tercero")
