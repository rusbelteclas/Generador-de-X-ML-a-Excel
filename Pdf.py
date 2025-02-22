import pandas as pd
import xml.etree.ElementTree as ET
import os

# Definir la ruta donde están los archivos XML
carpeta_facturas = r'C:\Users\cpjav\Downloads\COHETZALA\2025-01\XML'

# Diccionario de mapeo para los códigos de Forma de Pago
forma_pago_map = {
    '01': 'Efectivo',
    '03': 'Transferencia electrónica de fondos',
    '99': 'Otros'
}

# Lista para almacenar los datos extraídos
facturas_data = []

# Recorrer todos los archivos XML en la carpeta
for archivo in os.listdir(carpeta_facturas):
    if archivo.endswith(".xml"):
        ruta_xml = os.path.join(carpeta_facturas, archivo)
        try:
            tree = ET.parse(ruta_xml)
            root = tree.getroot()

            # Espacio de nombres para la versión 4.0
            namespaces = {'cfdi': 'http://www.sat.gob.mx/cfd/4', 'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'}

            # Extraer la información relevante del XML
            emisor = root.find(".//cfdi:Emisor", namespaces=namespaces)
            rfc_emisor = emisor.attrib.get('Rfc', 'N/A') if emisor is not None else 'N/A'
            nombre_emisor = emisor.attrib.get('Nombre', 'N/A') if emisor is not None else 'N/A'
            
            fecha = root.attrib.get('Fecha', 'N/A')
            total = root.attrib.get('Total', 'N/A')
            
            # Extraer el UUID del TimbreFiscalDigital dentro del Complemento
            uuid_node = root.find(".//cfdi:Complemento/tfd:TimbreFiscalDigital", namespaces=namespaces)
            uuid = uuid_node.attrib.get('UUID', 'N/A') if uuid_node is not None else 'N/A'
            
            # Extraer la Forma de Pago y traducirla
            forma_pago_code = root.attrib.get('FormaPago', 'N/A')
            forma_pago = forma_pago_map.get(forma_pago_code, 'Desconocida')  # Si el código no está en el diccionario, ponemos 'Desconocida'

            # Extraer los conceptos
            conceptos = []
            for concepto in root.findall(".//cfdi:Conceptos/cfdi:Concepto", namespaces=namespaces):
                clave_prod_serv = concepto.attrib.get('ClaveProdServ', 'N/A')
                cantidad = concepto.attrib.get('Cantidad', 'N/A')
                valor_unitario = concepto.attrib.get('ValorUnitario', 'N/A')
                importe = concepto.attrib.get('Importe', 'N/A')
                descripcion = concepto.attrib.get('Descripcion', 'N/A')
                
                # Crear una representación estructurada del concepto
                concepto_data = f"ClaveProdServ: {clave_prod_serv} | Cantidad: {cantidad} | ValorUnitario: {valor_unitario} | Importe: {importe} | Descripción: {descripcion}"
                conceptos.append(concepto_data)
            
            # Agregar los datos de la factura y los conceptos a la lista
            facturas_data.append({
                'RFC Emisor': rfc_emisor,
                'Nombre Emisor': nombre_emisor,
                'Fecha': fecha,
                'Total': total,
                'UUID': uuid,
                'Forma de Pago': forma_pago,
                'Conceptos': "; ".join(conceptos)  # Concatenar los conceptos en una sola celda
            })
        
        except ET.ParseError as e:
            print(f"Error al analizar el archivo {archivo}: {e}")
        except Exception as e:
            print(f"Error inesperado al procesar el archivo {archivo}: {e}")

# Crear un DataFrame de pandas
df = pd.DataFrame(facturas_data)

# Guardar el DataFrame como un archivo Excel
df.to_excel('facturas_conceptos_extraidos.xlsx', index=False)

print("¡El archivo Excel se ha generado exitosamente:)")
