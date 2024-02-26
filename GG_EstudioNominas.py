import pandas as pd

# Cargar los archivos Excel
archivo_aux = 'new.xlsx'
archivo_nomina = 'NOMINA.xlsx'

# Leer los archivos Excel
df_aux = pd.read_excel(archivo_aux)
df_nomina = pd.read_excel(archivo_nomina)

# Convertir las columnas de código a tipo de dato común
df_aux['Código Aleixos'] = df_aux['Código Aleixos'].astype(str)
df_nomina['CODIGO'] = df_nomina['CODIGO'].astype(str).str.strip('0')

# Realizar la unión de los dos dataframes por el campo de código
df_merged = pd.merge(df_nomina, df_aux, left_on='CODIGO', right_on='Código Aleixos', how='inner')

columnas_requeridas = ['Nombre', 'CODIGO',	'Salario Base',	'P.P.P.Marzo (Beneficios)',	'P.P.EXTRA','C.LINEAL','Departamento', 'Categoria puesto', 'Centro de Coste', 'CECO', 'Contrato','Categoría', 'Nivel Salarial', 'Grupo profesional', 'Antigüedad contrato']

df_resultado = df_merged[columnas_requeridas]

# Guardar el resultado en un nuevo archivo Excel
nombre_archivo_resultado = 'Estudio_Salarial_1.xlsx'
df_resultado.to_excel(nombre_archivo_resultado, index=False)

print(f"Se ha creado el archivo {nombre_archivo_resultado} con el resultado.")
