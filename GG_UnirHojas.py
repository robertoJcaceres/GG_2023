import calendar
import pandas as pd

# Nombres de los meses en español
meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
         'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# Ruta del archivo de Excel
excel_file = r'C:\Users\rcaceres\Documents\Costes mensuales.xlsx'

# Crea un DataFrame vacío para almacenar los datos combinados
df_combined = pd.DataFrame()

# Combina los datos de todas las hojas
for year in range(2019, 2024):
    for month in range(1, 13):
        month_name = meses[month - 1]  # Restamos 1 para obtener el índice correcto en la lista
        sheet_name = f'{month_name} {str(year)[2:]}'

        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df_combined = pd.concat([df_combined, df], ignore_index=True)
        except:
            # Si la hoja no existe, pasa al siguiente mes/año
            continue

# Guarda el DataFrame combinado en un nuevo archivo de Excel
output_file = 'tabla_salarios.xlsx'
df_combined.to_excel(output_file, index=False)
