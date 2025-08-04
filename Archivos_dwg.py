import os
from openpyxl import Workbook

# Ruta de la carpeta a revisar
ruta = r"\\auscorp.ausenco.com.au\Lima\Proj\102456-05\04 Drawing\04.03 Earthworks\AAE\02 Xref"

# Nombre del archivo Excel de salida
archivo_excel = "archivos_dwg.xlsx"

# Lista para almacenar los nombres de los archivos .dwg encontrados
archivos_dwg = []

# Recorrer la ruta y subcarpetas
for root, dirs, files in os.walk(ruta):
    for file in files:
        if file.lower().endswith('.dwg'):
            archivos_dwg.append(file)

# Crear un libro de Excel y una hoja
wb = Workbook()
ws = wb.active
ws.title = "Archivos DWG"

# Escribir los nombres de los archivos en la columna A
for i, nombre in enumerate(archivos_dwg, start=1):
    ws.cell(row=i, column=1, value=nombre)

# Guardar el archivo Excel
wb.save(archivo_excel)

print(f"Se encontraron {len(archivos_dwg)} archivos .dwg. El listado fue guardado en '{archivo_excel}'.")