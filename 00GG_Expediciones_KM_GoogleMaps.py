import pandas as pd
import googlemaps

# Reemplaza con tu clave de API de Google Maps
gmaps = googlemaps.Client(key='AIzaSyAKiy69olCymISDR8l6vFFHZcTpbHFN5K0')

# Función para obtener la ubicación (latitud, longitud) de una dirección
def get_location(address):
    try:
        geocode_result = gmaps.geocode(address)
        if geocode_result:
            location = geocode_result[0]['geometry']['location']
            return (location['lat'], location['lng'])
        else:
            return None, None
    except Exception as e:
        print(f"Error al obtener la ubicación de {address}: {e}")
        return None, None

# Función para calcular la distancia entre dos puntos
def calculate_distance(origen, destino):
    if origen and destino:
        try:
            directions_result = gmaps.directions(origen, destino, mode="driving")
            distance = directions_result[0]['legs'][0]['distance']['value'] / 1000  # Convertir de metros a kilómetros
            return distance
        except Exception as e:
            print(f"Error al calcular la distancia: {e}")
            return None
    else:
        return None

# Carga el archivo Excel
input_file_path = 'C:\\Users\\rcaceres\\OneDrive - Hinojosa Packaging Group\\GG_KM_Expedidos.xlsx'
df = pd.read_excel(input_file_path)

# Ingresa el año que deseas analizar
año_deseado = int(input("Ingresa el año deseado: "))

# Filtra el DataFrame por el año deseado
df_filtrado = df[df['Año'] == año_deseado]

# Extrae los valores únicos de clientes y direcciones para ese año
clientes_unicos = df_filtrado[['Descripcion_Cliente', 'Direccion']].drop_duplicates()

# Dirección de origen
origen_address = "Galeria Gràfica SL, Pol. Ind. Mas del Jutge, C/ Tonellet, 84-86, 46901 Torrent, Valencia, España"
origen_lat, origen_lng = get_location(origen_address)

# Calcula la distancia para cada cliente/dirección única y crea un diccionario
distancias_km = {}
for _, row in clientes_unicos.iterrows():
    cliente = row['Descripcion_Cliente']
    direccion = row['Direccion']
    lat, lng = get_location(direccion)
    if lat and lng:
        distancia = calculate_distance(f"{origen_lat},{origen_lng}", f"{lat},{lng}")
        distancias_km[(cliente, direccion)] = distancia
    else:
        distancias_km[(cliente, direccion)] = None

# Añade la distancia calculada a cada fila en el DataFrame filtrado
df_filtrado['km'] = df_filtrado.apply(lambda x: distancias_km.get((x['Descripcion_Cliente'], x['Direccion'])), axis=1)

# Guarda el DataFrame modificado en un nuevo archivo Excel
output_file_path = 'C:\\Users\\rcaceres\\Desktop\\Resultados_Distancias.xlsx'
df_filtrado.to_excel(output_file_path, index=False)

print(f"Archivo guardado en {output_file_path}")
