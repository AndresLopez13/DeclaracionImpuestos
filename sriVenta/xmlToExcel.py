import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook


def clean_xml_files(folder_path):
    new_folder_path = os.path.join(folder_path, 'cleaned')
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)
    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            with open(os.path.join(folder_path, filename), 'r', encoding='UTF-8') as file:
                xml_data = file.read()
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="UTF-8" standalone="no"?>', '')
                xml_data = xml_data.replace(
                    '<comprobante><![CDATA[<?xml version="1.0" encoding="UTF-8"?>', '')
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


def save_excel_file(folder_name, folder_path):
    file_name = 'Resumen declaracion Ventas ' + folder_name
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
    columna_sumas = 6

    for i in range(columna_sumas, len(factura)):
        if factura[i] == '-':
            factura[i] = '0'
        if i - columna_sumas < len(sumatorias):
            sumatorias[i - columna_sumas] += float(factura[i])

    return sumatorias


if __name__ == '__main__':
    path_name = 'marzo'
    folder_path = r'C:/Users/Andres/Documents/USB/DeclaracionImpuestos/sriVenta/2024/' + path_name
    folder_path_retenciones = folder_path + '/retenciones'
    folder_name = os.path.basename(folder_path)

    clean_xml_files(folder_path_retenciones)
    folder_path_retenciones = folder_path_retenciones + '/cleaned'

    workbook = Workbook()
    hoja_activa = workbook.active
    sumatorias = [0, 0, 0, 0, 0]
    contador = 1

    hoja_activa.append(['Num', 'Archivo', '# Fac.', 'Razon Social', 'Descripción', 'Fecha Emisión', 'Total Sin IVA',
                        'IVA', 'Importe Total', 'Ret. Renta', 'Ret. IVA', 'Observación'])

    archivos = [entry.name for entry in os.scandir(
        folder_path) if entry.is_file() and entry.name.endswith('.xml')]

    for archivo in archivos:
        # Ventas
        ruta_archivo = os.path.join(folder_path, archivo)
        try:
            tree = ET.parse(ruta_archivo)
        except ET.ParseError as e:
            print(
                f'Error: no se pudo procesar el archivo {archivo} debido a un error de sintaxis: {e}')
            continue
        rootVentas = tree.getroot()

        factura = []

        factura.append(contador)
        contador += 1
        factura.append(archivo)

        secuencial = int(rootVentas.find('.//secuencial').text)
        factura.append(secuencial)
        factura.append(rootVentas.find('.//razonSocialComprador').text)

        detalles = rootVentas.find('.//detalles')
        for detalle in detalles.findall('.//detalle'):
            descripcion = detalle.find('.//descripcion')
            factura.append(descripcion.text)
            break  # Se detiene el ciclo después de encontrar la primera descripción

        factura.append(rootVentas.find('.//fechaEmision').text)
        factura.append(float(rootVentas.find('.//totalSinImpuestos').text))

        totalConImpuestos = rootVentas.find('.//totalConImpuestos')
        for impuesto in totalConImpuestos.iter('totalImpuesto'):
            factura.append(float(impuesto.find('.//valor').text))

        factura.append(float(rootVentas.find('.//importeTotal').text))

        # Retenciones
        codigo1 = 0
        codigo2 = 0
        archivos_retenciones = [entry.name for entry in os.scandir(
            folder_path_retenciones) if entry.is_file() and entry.name.endswith('.xml')]

        for archivo_retencion in archivos_retenciones:
            ruta_archivo_retenciones = os.path.join(
                folder_path_retenciones, archivo_retencion)
            try:
                tree = ET.parse(ruta_archivo_retenciones)
            except ET.ParseError as e:
                print(
                    f'Error: no se pudo procesar el archivo {archivo_retencion} debido a un error de sintaxis: {e}')
                continue
            rootRetencion = tree.getroot()

            numDocSustento = rootRetencion.find('.//numDocSustento').text
            numDocSustento = int(numDocSustento[-9:])

            if numDocSustento == secuencial: # TODO agregar caso donde no se encuentre el numDocSustento y en su lugar agregue un guion al Excel
                # caso retencion con etiqueta docSustento y retencion con etiqueta impuesto TEST
                if rootRetencion.find('.//retencion') is not None:
                    etiquetaRetencion = 'retencion'
                else: 
                    etiquetaRetencion = 'impuesto'

                for retencion in rootRetencion.iter(etiquetaRetencion):
                    codigo = int(retencion.find('.//codigo').text)
                    valorRetenido = float(
                        retencion.find('.//valorRetenido').text)

                    if codigo == 1:
                        codigo1 += valorRetenido
                    elif codigo == 2:
                        codigo2 += valorRetenido
                    else:
                        print(f'Existe otro codigo de retencion: {codigo}')

        factura.append(codigo1) 
        factura.append(codigo2)
        hoja_activa.append(factura)
        sumatorias = add_totals(factura) 

    hoja_activa.append(['TOTAL', '', '', '', '', '', *sumatorias])

    save_excel_file(folder_name, folder_path)
