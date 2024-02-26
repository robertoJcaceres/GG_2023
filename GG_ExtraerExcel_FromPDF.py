import camelot
import pandas as pd

# Ruta del archivo PDF
pdf_path = "C:\\Users\\rcaceres\\Desktop\\GGrafica-Coding\\GGrafica\\CIRCULAR 10 2023 ANEXO 2 TABLAS SALARIALES 2023 ACTUALIZADAS CON ANTIGUEDAD (002).pdf"

# Extraer las tablas del PDF
tablas_pdf = camelot.read_pdf(pdf_path, pages='all')

# Crear un archivo de Excel
excel_path = "ruta/del/archivo.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')

# Guardar cada tabla en una hoja de Excel
for i, tabla in enumerate(tablas_pdf):
    df = tabla.df
    df.to_excel(writer, sheet_name=f'Hoja{i+1}', index=False)

# Guardar y cerrar el archivo de Excel
writer.save()
writer.close()
