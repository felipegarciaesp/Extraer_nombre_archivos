import os
from openpyxl import Workbook

ruta = r"\\auscorp.ausenco.com.au\Lima\Proj\102456-05\04 Drawing\04.03 Earthworks\AAE\02 Xref"
archivo_excel = "archivos_dwg.xlsx"

# Solo listar los archivos en el directorio principal (sin subdirectorios)
archivos_dwg = [
    f for f in os.listdir(ruta)
    if os.path.isfile(os.path.join(ruta, f)) and f.lower().endswith('.dwg')
]

wb = Workbook()
ws = wb.active
ws.title = "Archivos DWG"

for i, nombre in enumerate(archivos_dwg, start=1):
    ws.cell(row=i, column=1, value=nombre)

wb.save(archivo_excel)
print(f"Se encontraron {len(archivos_dwg)} archivos .dwg en la carpeta principal. El listado fue guardado en '{archivo_excel}'.")