import pandas as pd
import os
from datetime import datetime

def extraer_fecha(nombre_archivo):
    """
    Intenta extraer la fecha de un nombre de archivo en varios formatos.
    """
    fecha_str = "_".join(nombre_archivo.split('_')[-3:]).replace('.xlsx', '')
    formatos_fecha = ['%d_%m_%Y', '%d-%m-%Y']  # Añadir aquí otros formatos de fecha si son necesarios

    for formato in formatos_fecha:
        try:
            return datetime.strptime(fecha_str, formato).date()
        except ValueError:
            continue

    # Si llegamos a este punto, ninguno de los formatos funcionó
    raise ValueError(f"Formato de fecha no reconocido en el archivo {nombre_archivo}.")

def main():
    directorio = r'C:\Users\rcaceres\OneDrive - Hinojosa Packaging Group\Inventarios'  # Directorio actualizado
    archivos = [f for f in os.listdir(directorio) if f.endswith('.xlsx')]
    df_list = []

    # ... (resto del script sigue igual)


    for archivo in archivos:
        try:
            fecha = extraer_fecha(archivo)

            # Leer el archivo en un dataframe
            df = pd.read_excel(os.path.join(directorio, archivo))
            
            # Agregar la columna de fecha
            df['Fecha de Inventario'] = fecha
            
            # Agregar el dataframe a la lista
            df_list.append(df)

        except Exception as e:
            print(f"Error procesando {archivo}: {e}")
            continue  # Continuar con el siguiente archivo si hay un error

    if df_list:
        # Concatenar todos los dataframes en uno solo
        df_final = pd.concat(df_list, ignore_index=True)
        
        # Guardar en un nuevo archivo Excel
        ruta_archivo_salida = os.path.join(directorio, 'inventario_concatenado.xlsx')
        df_final.to_excel(ruta_archivo_salida, index=False)
        print(f"Inventario consolidado guardado en {ruta_archivo_salida}")
    else:
        print("No se pudo crear un inventario consolidado debido a errores en los archivos de entrada.")

# Punto de entrada principal
if __name__ == "__main__":
    main()
