import pytesseract
import subprocess
from PIL import ImageGrab
import pyautogui
import time
import re
import json
import pandas as pd
import os
import shutil
import pdfplumber  # Biblioteca para extraer texto de PDFs
import win32com.client as win32  # Para enviar correos con Outlook
from pathlib import Path


# Retrocede un nivel desde la carpeta de scripts
# Usar Path en lugar de os.path para la ruta raíz
ruta_raiz = Path(__file__).parent.parent
config_path = ruta_raiz / "config" / "config.json"
with open(config_path, 'r') as file:
    config = json.load(file)

# Configuración de pytesseract
pytesseract.pytesseract.tesseract_cmd = config["paths"]["tesseract"]

# Rutas principales
carpeta = ruta_raiz / config["paths"]["inputs"]
carpeta_descargas = ruta_raiz / config["paths"]["downloads"]
path_pdf = ruta_raiz / config["paths"]["pdf"]
config_folder = ruta_raiz / config["paths"]["config"]
output_folder = ruta_raiz / config["paths"]["output"]
origen = ruta_raiz / config["paths"]["origen_folder"]
destino = carpeta
# Obtener la lista de documentos a excluir
documentos_excluir = config.get("tipo_documento_excluir", [])

# Validación de rutas (opcional)
if config["validation"]["check_paths"]:
    for key, path in config["paths"].items():
        if not os.path.exists(path):
            print(f"Advertencia: La ruta {path} no existe.")

# -------------------------------
# CONFIGURACIÓN DE EXCEL
# -------------------------------

columna_a_iterar = "CUFE/CUDE"
columna_nit_receptor = "NIT Receptor"
columna_procesado = "PDF Almacenado"  # Columna para marcar si ya se procesó
# Nueva columna para almacenar información del PDF
columna_info_pdf = "Información PDF"


# -------------------------------
# ----- CONFIGURACIÓN INICIAL ---#
# -------------------------------

# Contador de ejecuciones
max_ejecuciones = 3
ejecuciones_realizadas = 0


# Función para enviar correo electrónico


def enviar_correo(archivo):
    try:
        # 📩 Crear una instancia de Outlook
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0 representa un correo nuevo

        # ✉ Configurar el correo
        mail.Subject = "Notificación: Ejecución del Bot Finalizada"
        mail.Body = "El bot ha finalizado su ejecución. Todas las filas han sido procesadas o se alcanzó el límite de 3 intentos."
        mail.To = "jeferson.vargara@acafi.com.co"
        ruta_archivo = fr"C:\Users\santi\OneDrive\Escritorio\Swith_bots\Swith_bots\ACAFI\inputs\{archivo}.xlsx"
        # Verificar si el archivo existe
        if os.path.exists(ruta_archivo):
            print(f"✅ El archivo existe en: {ruta_archivo}")

            # Verificar si el archivo es accesible
            try:
                with open(ruta_archivo, 'rb') as f:
                    print("✅ El archivo es accesible para lectura.")
                    mail.Attachments.Add(ruta_archivo)
                    print("Archivo adjuntado correctamente.")
            except Exception as e:
                print(f"❌ No se pudo acceder al archivo: {e}")
        else:
            print(f"⚠ Archivo no encontrado: {ruta_archivo}")

        # 📤 Enviar el correo
        mail.Send()
        print("✅ Correo electrónico enviado correctamente.")
    except Exception as e:
        print(f"❌ Error al enviar el correo electrónico: {e}")


def mover_excels(origen, destino):
    """
    Mueve todos los archivos Excel de una carpeta origen a una carpeta destino.

    Parámetros:
        origen (str): Ruta de la carpeta donde se buscarán los archivos Excel.
        destino (str): Ruta de la carpeta donde se moverán los archivos.

    Retorna:
        tuple: (archivos_movidos, mensaje)
    """
    # Validar si la carpeta origen existe
    if not os.path.exists(origen):
        return [], f"Error: La carpeta origen '{origen}' no existe."

    # Crear carpeta destino si no existe
    os.makedirs(destino, exist_ok=True)

    # Buscar archivos Excel (.xlsx y .xls)

    archivos_excel = [
        f for f in os.listdir(origen)
        if f.endswith(('.xlsx', '.xls'))
    ]

    # Mover archivos
    archivos_movidos = []
    for archivo in archivos_excel:
        origen_path = os.path.join(origen, archivo)
        destino_path = os.path.join(destino, archivo)

        # Mover y evitar sobrescribir
        if os.path.exists(destino_path):
            base, ext = os.path.splitext(archivo)
            nuevo_nombre = f"{base}_DUPLICADO{ext}"
            destino_path = os.path.join(destino, nuevo_nombre)

        shutil.move(origen_path, destino_path)
        archivos_movidos.append(archivo)

    # Mensaje de resultado
    if not archivos_excel:
        return [], f"No se encontraron archivos Excel en '{origen}'."
    else:
        return archivos_movidos, f"Se movieron {len(archivos_excel)} archivos a '{destino}'."


archivos, mensaje = mover_excels(origen, destino)
print(mensaje)
# Bucle principal del bot
while ejecuciones_realizadas < max_ejecuciones:
    ejecuciones_realizadas += 1
    print(f"\nEjecución número: {ejecuciones_realizadas}")

    if archivos:
        print("Archivos movidos:", ", ".join(archivos))
    try:
        # Recorrer todos los archivos en la carpeta de entrada
        for archivo in os.listdir(carpeta):
            if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
                # Construir la ruta completa del archivo
                ruta_archivo = os.path.join(carpeta, archivo)
                ruta_excel_json = os.path.join(
                    config_folder, "ruta_excel.json")
                nombre_archivo = ruta_archivo.split("\\")[-1]
                # Cargar el archivo Excel
                df = pd.read_excel(ruta_archivo, engine="openpyxl")
                # Extrae el NIT (todo antes del primer '(', '_' o '.')
                nit_receptor = re.split(r'[_(.]', nombre_archivo)[0]
                # cargar el archivo excel que contiene la base de datos con los codigos de producto
                bd_terceros_path = os.path.join(
                    config_folder, f"{nit_receptor}.xlsx")
                df2 = pd.read_excel(bd_terceros_path)
                # Si las columnas "Procesado" e "Información PDF" no existen, las creamos
                if columna_procesado not in df.columns:
                    # Por defecto, marcamos como no procesado
                    df[columna_procesado] = "No"
                if columna_info_pdf not in df.columns:
                    # Columna vacía para la información del PDF
                    df[columna_info_pdf] = ""

                # Verificar si todas las filas ya están procesadas
                if all(df[columna_procesado] == "Sí"):
                    print(
                        f"Todas las filas del archivo {archivo} ya están procesadas. Saltando...")
                    continue  # Saltar este archivo y continuar con el siguiente

                # Filtrar filas que NO estén en la lista de exclusión
                df_filtrado = df[~df["Tipo de documento"].astype(
                    str).str.strip().isin(documentos_excluir)]

                # Sobrescribir el archivo original con el filtrado sin índice
                df_filtrado.to_excel(
                    ruta_archivo, index=False, engine="openpyxl")

                # Guardar la ruta en un JSON
                ruta_archivo_json = {"ruta_archivo.excel": ruta_archivo}
                with open(ruta_excel_json, 'w', encoding='utf-8') as f:
                    json.dump(ruta_archivo_json, f,
                              ensure_ascii=False, indent=4)

                # Verificar si el archivo ya está completamente procesado
                if all(df[columna_procesado] == "Sí"):
                    print(
                        f"El archivo {archivo} ya está completamente procesado. Saltando...")
                    continue  # Saltar este archivo y continuar con el siguiente

                # Paso 1: Verificar y crear las columnas si no existen
                columnas_requeridas = ['Nombre del producto',
                                       'codigo de producto', 'centro de costos']

                for columna in columnas_requeridas:
                    if columna not in df.columns:
                        df[columna] = ''  # Crear la columna con valores vacíos

                # Paso 1: Convertir las columnas relevantes a tipo str
                df['Nombre del producto'] = df['Nombre del producto'].astype(
                    str)
                df['codigo de producto'] = df['codigo de producto'].astype(str)
                df['centro de costos'] = df['centro de costos'].astype(str)

                # Convertir columnas relevantes de df2 a tipo str
                df2['Nit emisor'] = df2['Nit emisor'].astype(str)
                df2['Nombre del producto'] = df2['Nombre del producto'].astype(
                    str)
                df2['Código del Producto'] = df2['Código del Producto'].fillna(
                    0).astype(float).astype(int).astype(str)  # Manejar NaN
                df2['Centro de Costo'] = df2['Centro de Costo'].astype(str)

                # Paso 2: Recorrer cada fila del archivo principal (df)
                for indice, fila in df.iterrows():
                    # Obtener el NIT Emisor de la fila actual
                    nit_amisor = fila['NIT Emisor']
                    bd_nit_emisor = df2['Nit emisor'].astype(
                        str)  # Asegurar que el NIT en df2 sea str

                    # Buscar coincidencias en el archivo de búsqueda (df2)
                    coincidencias = df2[bd_nit_emisor == str(nit_amisor)]

                    # Si hay coincidencias, extraer los datos necesarios y actualizar las columnas en df
                    if not coincidencias.empty:
                        # Tomar la primera coincidencia
                        primera_coincidencia = coincidencias.iloc[0]
                        df.at[indice, 'Nombre del producto'] = primera_coincidencia['Nombre del producto']
                        # Ya es str
                        df.at[indice, 'codigo de producto'] = primera_coincidencia['Código del Producto']
                        df.at[indice, 'centro de costos'] = primera_coincidencia['Centro de Costo']
                    else:
                        # Si no hay coincidencias, asignar "sin coincidencia" y dejar las otras columnas en blanco
                        df.at[indice, 'Nombre del producto'] = 'sin coincidencia'
                        df.at[indice, 'codigo de producto'] = ''  # Ya es str
                        df.at[indice, 'centro de costos'] = ''

                # Paso 3: Verificar si la columna "CUFE/CUDE" existe
                if 'CUFE/CUDE' in df.columns:
                    print("La columna 'CUFE/CUDE' existe en el DataFrame.")
                else:
                    print("La columna 'CUFE/CUDE' no existe en el DataFrame.")

                # Paso 4: Guardar el DataFrame actualizado en un archivo Excel
                # Usar index=False para evitar guardar el índice
                df.to_excel(ruta_archivo, index=False)
                # Verificar si la columna "CUFE/CUDE" existe en el archivo
                if columna_a_iterar in df.columns:
                    # Iniciar la aplicación que se usará para la automatización
                    app_id = "shell:AppsFolder\\57778KONTALID.KONTALIDTools_1crwx9b2rpxma!com.embarcadero.KONTALIDTools"
                    process = subprocess.Popen(
                        ["explorer.exe", app_id], shell=True)

                    # Esperar a que la aplicación se inicie
                    time.sleep(5)

                    # Capturar la pantalla y extraer texto con OCR
                    screenshot = ImageGrab.grab()
                    text = pytesseract.image_to_string(screenshot)

                    # Buscar un texto específico en la pantalla
                    search_text = "Documento"
                    if search_text in text:
                        print(
                            "Texto encontrado. Haciendo clic en el área correspondiente...")
                        data = pytesseract.image_to_data(
                            screenshot, output_type=pytesseract.Output.DICT)

                        for i, word in enumerate(data['text']):
                            if word.strip() == search_text:
                                x1, y1, x2, y2 = data['left'][i], data['top'][i], data['left'][i] + \
                                    data['width'][i], data['top'][i] + \
                                    data['height'][i]
                                click_x = (x1 + x2) // 2
                                click_y = (y1 + y2) // 2
                                time.sleep(2)
                                pyautogui.click(click_x, click_y)
                                print(
                                    f"Haciendo clic en ({click_x}, {click_y})")
                                time.sleep(3)
                                pyautogui.press('tab')

                    # Iterar sobre los valores de la columna del Excel
                    for index, row in df.iterrows():
                        valor = row[columna_a_iterar]
                        procesado = row[columna_procesado]

                        # Si ya está procesado, lo saltamos
                        if procesado == "Sí":
                            print(
                                f"El CUFE {valor} ya fue procesado. Saltando...")
                            continue

                        # Escribir el CUFE en la aplicación
                        time.sleep(2)
                        pyautogui.write(valor)
                        time.sleep(1)
                        pyautogui.press('enter')
                        time.sleep(5)

                        # Buscar el botón "Descargar" en pantalla
                        search_text = "Descargar"
                        text_found = False
                        timeout = 30
                        start_time = time.time()

                        while not text_found:
                            screenshot = ImageGrab.grab()
                            text = pytesseract.image_to_string(screenshot)
                            if search_text in text:
                                text_found = True
                                print(f"Texto '{search_text}' encontrado.")
                            elif time.time() - start_time > timeout:
                                print(
                                    "Se alcanzó el tiempo de espera máximo sin encontrar el texto.")
                                break
                            else:
                                print("Texto no encontrado, esperando...")
                                time.sleep(1)

                        if text_found:
                            data = pytesseract.image_to_data(
                                screenshot, output_type=pytesseract.Output.DICT)
                            for i, word in enumerate(data['text']):
                                if word.strip() == search_text:
                                    time.sleep(2)
                                    x1, y1 = data['left'][i], data['top'][i]
                                    x2, y2 = x1 + \
                                        data['width'][i], y1 + \
                                        data['height'][i]
                                    click_x = (x1 + x2) // 2
                                    click_y = (y1 + y2) // 2
                                    time.sleep(1)
                                    pyautogui.click(click_x, click_y)
                                    print(
                                        f"Haciendo clic en ({click_x}, {click_y})")
                                    time.sleep(2)
                                    break
                        else:
                            df.at[index, columna_procesado] = "No"
                            break

                        # Esperar la descarga del archivo PDF
                        tiempo_max_espera = 60
                        tiempo_transcurrido = 0
                        intervalo_espera = 2
                        archivo_descargado = None
                        intentos = 0
                        while tiempo_transcurrido < tiempo_max_espera:
                            archivos = [os.path.join(carpeta_descargas, archivo) for archivo in os.listdir(
                                carpeta_descargas) if archivo.endswith(".pdf")]
                            if archivos:
                                archivo_descargado = max(
                                    archivos, key=os.path.getmtime)
                                break
                            time.sleep(intervalo_espera)
                            tiempo_transcurrido += intervalo_espera
                        time.sleep(2)

                        # Mover el archivo descargado a la carpeta destino
                        if archivo_descargado:
                            nombre_archivo = os.path.basename(
                                archivo_descargado)
                            # Ruta de la carpeta que quieres verificar
                            ruta_carpeta = os.path.join(
                                path_pdf, str(nit_receptor))

                            # Verificar si la carpeta existe
                            if not os.path.exists(ruta_carpeta):
                                # Si no existe, crearla
                                os.makedirs(ruta_carpeta)
                                print(
                                    f"La carpeta '{ruta_carpeta}' ha sido creada.")
                            else:
                                print(
                                    f"La carpeta '{ruta_carpeta}' ya existe.")

                            destino_final = os.path.join(
                                ruta_carpeta, nombre_archivo)
                            os.makedirs(ruta_carpeta, exist_ok=True)

                            # Mover el archivo a la nueva ubicación
                            shutil.move(archivo_descargado, destino_final)

                            # Verificar si el archivo se ha movido correctamente
                            if os.path.exists(destino_final):
                                print(
                                    f"Archivo '{nombre_archivo}' movido exitosamente a: {destino_final}")
                                # Marcar como procesado
                                df.at[index, columna_procesado] = "Sí"
                            else:
                                print(
                                    f"Error: El archivo '{nombre_archivo}' no se ha movido correctamente.")
                                df.at[index, columna_procesado] = "No"
                                # No marcar como procesado si no se movió correctamente
                        else:
                            print(
                                "No se encontró ningún archivo PDF dentro del tiempo esperado.")
                            continue  # Continuar con la siguiente iteración

                        # Extraer información del PDF
                        try:
                            with pdfplumber.open(destino_final) as pdf:
                                descripcion = ""  # Variable para almacenar la descripción encontrada
                                for page in pdf.pages:
                                    tables = page.extract_tables()  # Extraer todas las tablas de la página
                                    for table in tables:
                                        # Verificar que la tabla no esté vacía y tenga al menos 2 filas
                                        if table and len(table) > 1:
                                            # Buscar la columna que contiene "descri" en los encabezados (segunda fila)
                                            encabezados = table[1]
                                            columna_descripcion = None
                                            for i, encabezado in enumerate(encabezados):
                                                if encabezado and "descri" in encabezado.lower():
                                                    columna_descripcion = i
                                                    print(
                                                        f"Columna 'Descripción' encontrada en el índice: {columna_descripcion}")
                                                    break

                                            # Si se encontró la columna, buscar el valor en las filas siguientes
                                            if columna_descripcion is not None:
                                                # Ignorar las filas de encabezados
                                                for row in table[2:]:
                                                    if len(row) > columna_descripcion and row[columna_descripcion]:
                                                        # Unir las líneas de la descripción si está dividida
                                                        descripcion = " ".join(
                                                            str(row[columna_descripcion]).split("\n"))
                                                        print(
                                                            f"Descripción extraída: {descripcion}")
                                                        break
                                                if descripcion:
                                                    break  # Salir del bucle si se encontró la descripción
                                        if descripcion:
                                            break  # Salir del bucle de páginas si se encontró la descripción

                                # Mostrar la descripción encontrada
                                if descripcion:
                                    print(
                                        f"Descripción encontrada: {descripcion}")
                                else:
                                    print(
                                        "No se encontró la columna 'Descripción' o variantes en el PDF.")

                                # Guardar la información en la nueva columna
                                df.at[index, columna_info_pdf] = descripcion if descripcion else "Descripción no encontrada"
                        except Exception as e:
                            print(f"Error al extraer información del PDF: {e}")
                            df.at[index, columna_info_pdf] = "Error al extraer información"

                        # Guardar el DataFrame actualizado en el archivo Excel
                        df.to_excel(ruta_archivo, index=False)

                        # Salir del modo de descarga y limpiar la barra de búsqueda
                        pyautogui.press('esc')
                        time.sleep(2)
                        pyautogui.hotkey('ctrl', 'a')
                        time.sleep(2)
                        # Simula la pulsación de la tecla Delete
                        pyautogui.press('delete')

                    # Cerrar la aplicación después de procesar el archivo
                    app_name = "KONTALIDTools.exe"  # Reemplaza con el nombre real del ejecutable
                    subprocess.run(
                        ["taskkill", "/f", "/im", app_name], shell=True)
                    print(
                        f"La aplicación se ha cerrado después de procesar el archivo: {archivo}")
                else:
                    print(
                        f"La columna '{columna_a_iterar}' no existe en el archivo.")
            else:
                print("el archivo no es un documento excel")
                continue
        # Enviar correo electrónico al finalizar
        if all(df[columna_procesado] == "Sí") or ejecuciones_realizadas == 3:
            enviar_correo(nit_receptor)
        # Verificar si todas las filas están procesadas después de cada ejecución
        if all(df[columna_procesado] == "Sí"):
            print(
                "Todas las filas han sido procesadas. Deteniendo el bot antes de completar las 3 ejecuciones.")
            break

    except Exception as e:
        print(f"Ocurrió un error al procesar el archivo: {e}")
        # Cerrar la aplicación después de procesar el archivo
        app_name = "KONTALIDTools.exe"  # Reemplaza con el nombre real del ejecutable
        subprocess.run(
            ["taskkill", "/f", "/im", app_name], shell=True)
        print(
            f"La aplicación se ha cerrado después de procesar el archivo: {archivo}")

print("El bot ha finalizado.")
