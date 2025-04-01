import os
import json
import pdfplumber
import re


def extract_vendor_info(text):
    vendor_info = {
        "Tipo de contribuyente": None,
        "Departamento": None,
        "Régimen fiscal": None
    }

    vendor_section = re.search(
        r"Datos del (Emisor / Vendedor|vendedor)(.*?)(Datos del Adquiriente / Comprador|$)",
        text,
        re.DOTALL | re.IGNORECASE,
    )
    if vendor_section:
        vendor_text = vendor_section.group(2)
        lines = vendor_text.split("\n")
        for line in lines:
            if re.search(r"tipo\s*de\s*contribuyente", line, re.IGNORECASE):
                parts = re.split(r":\s*", line, maxsplit=1)
                if len(parts) > 1:
                    tipo_contribuyente = parts[1].strip()
                    if re.search(r"natural", tipo_contribuyente, re.IGNORECASE):
                        vendor_info["Tipo de contribuyente"] = "Persona Natural"
                    elif re.search(r"jur[ií]dica", tipo_contribuyente, re.IGNORECASE):
                        vendor_info["Tipo de contribuyente"] = "Persona Jurídica"
                    else:
                        vendor_info["Tipo de contribuyente"] = tipo_contribuyente

            if re.search(r"departamento", line, re.IGNORECASE):
                parts = re.split(r":\s*", line, maxsplit=1)
                if len(parts) > 1:
                    vendor_info["Departamento"] = parts[1].strip()

            if re.search(r"r[eé]g[ií]men\s*(f[ií]scal|tributario)?", line, re.IGNORECASE):
                parts = re.split(r":\s*", line, maxsplit=1)
                if len(parts) > 1:
                    regimen_fiscal = parts[1].strip()
                    match = re.search(
                        r"([A-Za-z]-\d+(-[A-Za-z]+)?)", regimen_fiscal)
                    if match:
                        vendor_info["Régimen fiscal"] = match.group(1)
                    else:
                        vendor_info["Régimen fiscal"] = regimen_fiscal
    return vendor_info


def extract_payment_method(text):
    match = re.search(r"Forma de pago:\s*(.*)\n", text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def extract_product_description(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and len(row) > 2:
                        if row[2] and "Descri" in row[2]:
                            description_index = table.index(row) + 1
                            if description_index < len(table):
                                return table[description_index][2]
    return None


def extract_total_bruto_factura(text):
    """
    Extrae el valor de "Total Bruto Factura" del texto del PDF.
    """
    total_bruto_match = re.search(
        r"Total Bruto Factura\s*([\d.,]+)", text, re.IGNORECASE
    )
    if total_bruto_match:
        return total_bruto_match.group(1).replace(".", "").replace(",", ".")
    return None


def process_pdf(pdf_file_path):
    with pdfplumber.open(pdf_file_path) as pdf:
        text = "".join(page.extract_text() or "" for page in pdf.pages)

    vendor_info = extract_vendor_info(text)
    payment_method = extract_payment_method(text)
    product_description = extract_product_description(pdf_file_path)
    total_bruto_factura = extract_total_bruto_factura(text)

    extracted_data = {
        "Archivo": pdf_file_path,
        "Información del vendedor": vendor_info,
        "Forma de Pago": payment_method,
        "Descripción del producto": product_description,
        "Total Bruto Factura": total_bruto_factura  # Nuevo campo agregado
    }
    return extracted_data


config_folder = r"C:\Users\santi\OneDrive\Escritorio\Swith_bots\Swith_bots\ACAFI\config"
json_path = os.path.join(config_folder, 'pdf_routes.json')
output_json_path = os.path.join(config_folder, 'datos_extraidos.json')

with open(json_path, 'r') as file:
    pdf_routes = json.load(file)

pdf_file_paths = pdf_routes.get('path_pdf', [])
if isinstance(pdf_file_paths, str):
    pdf_file_paths = [pdf_file_paths]

all_extracted_data = []
for pdf_file_path in pdf_file_paths:
    if os.path.exists(pdf_file_path):
        try:
            extracted_data = process_pdf(pdf_file_path)
            all_extracted_data.append(extracted_data)
        except Exception as e:
            print(f"Error al procesar el archivo {pdf_file_path}: {str(e)}")
    else:
        print(f"El archivo {pdf_file_path} no existe.")

with open(output_json_path, 'w', encoding='utf-8') as output_file:
    json.dump(all_extracted_data, output_file, ensure_ascii=False, indent=4)

print(f"Datos extraídos guardados en {output_json_path}")