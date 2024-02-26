import pandas as pd
from geopy.geocoders import Nominatim
from geopy.distance import geodesic

# Inicializar Nominatim API
geolocator = Nominatim(user_agent="myGeocoder")

# Definir la función para obtener coordenadas
def get_location(address):
    try:
        print(f"Obteniendo coordenadas para: {address}")
        location = geolocator.geocode(address)
        if location:
            return (location.latitude, location.longitude)
        else:
            print(f"No se encontraron coordenadas para: {address}")
            return (None, None)
    except Exception as e:
        print(f"Error al obtener coordenadas para: {address}. Error: {e}")
        return (None, None)

# Cargar datos del Excel
input_file_path = "C:\\Users\\rcaceres\\OneDrive - Hinojosa Packaging Group\\GG_KM_Expedidos.xlsx"
df = pd.read_excel(input_file_path)

# Filtrar por el año 2023
df_2023 = df[df['Año'] == 2023]

# Extraer valores únicos de clientes con sus direcciones
clientes_unicos = df_2023[['Descripcion_Cliente', 'Direccion']].drop_duplicates()

# Dirección de origen y obtener sus coordenadas
origen_address = "Galeria Gràfica SL, Pol. Ind. Mas del Jutge, C/ Tonellet, 84-86, 46901 Torrent, Valencia, España"
origen_coord = get_location(origen_address)

# Diccionario para almacenar la distancia por cliente/dirección
distancias = {}

# Calcular distancia para cada cliente/dirección única
for index, row in clientes_unicos.iterrows():
    cliente = row['Descripcion_Cliente']
    direccion = row['Direccion']
    print(f"Calculando distancia para el cliente: {cliente}")
    dest_coord = get_location(direccion)
    if None not in origen_coord and None not in dest_coord:
        distancia = geodesic(origen_coord, dest_coord).km
        distancias[(cliente, direccion)] = distancia
    else:
        distancias[(cliente, direccion)] = None

# Asignar distancia a cada albarán basándose en el cliente/dirección
df_2023['km'] = df_2023.apply(lambda x: distancias.get((x['Descripcion_Cliente'], x['Direccion'])), axis=1)

# Guardar a un nuevo archivo Excel
output_file_path = "C:\\Users\\rcaceres\\Desktop\\Distancias_2023.xlsx"
df_2023.to_excel(output_file_path, index=False)

print(f"Archivo guardado en {output_file_path}")
