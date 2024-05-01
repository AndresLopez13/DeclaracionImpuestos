import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook


def clean_xml_files(folder_path):
    new_folder_path = os.path.join(folder_path, 'cleaned')
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)
    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='UTF-8') as file:
                xml_data = file.read()
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="UTF-8" standalone="no"?>', '')
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="UTF-8"?>', '')
                # xml_data = xml_data.replace(
                #     "<?xml version='1.0' encoding='UTF-8'?>", '')
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="utf-8" standalone="no"?>', '')
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '')
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="utf-8" standalone="yes"?>', '')
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="utf-8"?>', '')
                xml_data = xml_data.replace(']]></comprobante>', '')
                xml_data = xml_data.replace('<comprobante><![CDATA[', '')
            with open(os.path.join(new_folder_path, filename), 'w', encoding='UTF-8') as file:
                file.write(xml_data)
            os.remove(file_path)


def save_excel_file(folder_name, folder_path):
    file_name = 'Resumen declaracion Compras ' + folder_name
    file_path = os.path.join(folder_path, file_name)

    if os.path.exists(file_path + '.xlsx'):
        i = 1
        while os.path.exists(file_path + f'_{i}.xlsx'):
            i += 1
        file_path = file_path + f'_{i}.xlsx'
    else:
        file_path = file_path + '.xlsx'

    workbook.save(file_path)
    print(f'Archivo generado: {file_path}')


def add_totals(factura):
    """ 
    Suma los valores de las facturas en las columnas de sumatorias
    :param factura: lista con los datos de la factura
    :return: lista con los valores de las sumatorias
    """
    columna_sumas = 6

    for i in range(columna_sumas, len(factura)):
        if i-columna_sumas < len(sumatorias):
            sumatorias[i-columna_sumas] += float(factura[i])

    return sumatorias


if __name__ == '__main__':
    path_name = 'marzo'
    folder_path = r'C:/Users/Andres/Documents/USB/DeclaracionImpuestos/sriCompra/2024/' + path_name
    clean_xml_files(folder_path)

    folder_name = os.path.basename(folder_path)

    workbook = Workbook()
    hoja_activa = workbook.active
    contador = 1
    sumatorias = [0, 0, 0]

    hoja_activa.append(['Num', 'Archivo', '# Fac.', 'Razon Social', 'Descripción', 'Fecha Emisión', 'Total Sin IVA',
                        'IVA', 'Importe Total'])

    folder_path_cleaned = folder_path + '/cleaned'
    archivos = [entry.name for entry in os.scandir(
        folder_path_cleaned) if entry.is_file() and entry.name.endswith('.xml')]

    for archivo in archivos:
        ruta_archivo = os.path.join(folder_path_cleaned, archivo)
        try:
            tree = ET.parse(ruta_archivo)
        except ET.ParseError as e:
            print(
                f'Error: no se pudo procesar el archivo {archivo} debido a un error de sintaxis: {e}')
            continue
        root = tree.getroot()

        factura = []

        factura.append(contador)
        contador += 1
        factura.append(archivo)

        secuencial = root.find('.//secuencial')
        factura.append(int(secuencial.text))

        razonSocial = root.find('.//razonSocial')
        factura.append(razonSocial.text)

        detalles = root.find('.//detalles')
        for detalle in detalles.findall('.//detalle'):
            descripcion = detalle.find('.//descripcion')
            factura.append(descripcion.text)
            break  # Se detiene el ciclo después de encontrar la primera descripción

        fechaEmision = root.find('.//fechaEmision')
        factura.append(fechaEmision.text)

        totalSinImpuestos = root.find('.//totalSinImpuestos')
        factura.append(float(totalSinImpuestos.text))

        totalConImpuestos = root.find('.//totalConImpuestos')
        suma_iva = 0
        for impuesto in totalConImpuestos.iter('totalImpuesto'):
            valor = impuesto.find('.//valor')
            suma_iva += float(valor.text)
            print(f'Suma iva: {suma_iva}')
        factura.append(suma_iva)

        importeTotal = root.find('.//importeTotal')
        factura.append(float(importeTotal.text))

        hoja_activa.append(factura)
        print("-----")
        sumatorias = add_totals(factura)

    hoja_activa.append(['TOTAL', '', '', '', '', '', *sumatorias])

    save_excel_file(folder_name, folder_path)
