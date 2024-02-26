import pandas as pd
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt

# Cargar el archivo csv
df = pd.read_csv('C:\\Users\\rcaceres\\Downloads\\data.csv', delimiter=',')

# Visualizar las primeras filas del DataFrame
print(df.head())

# Estadísticas descriptivas
print(df.describe())

# Suponiendo que df es tu DataFrame
# Convertir Margen Bruto a decimal
# Reemplaza las comas por puntos y elimina los espacios en blanco
df['Margen Bruto'] = df['Margen Bruto'].str.replace(',', '.').str.strip('% ').astype(float) / 100.0

# Estadísticas descriptivas
print(df[['Total Facturación', 'Margen Bruto']].describe())

# Visualización de datos
plt.figure(figsize=(10,5))

plt.subplot(1, 2, 1)
plt.hist(df['Total Facturación'], bins=30, edgecolor='k')
plt.title('Distribución de Total Facturación')

plt.subplot(1, 2, 2)
plt.hist(df['Margen Bruto'], bins=30, edgecolor='k')
plt.title('Distribución de Margen Bruto')

plt.show()

# Calcular percentiles
for percentile in [25, 50, 75, 90]:
    print(f'{percentile}th Percentile of Total Facturación: {df["Total Facturación"].quantile(percentile / 100):.2f}')
    print(f'{percentile}th Percentile of Margen Bruto: {df["Margen Bruto"].quantile(percentile / 100):.2f}')
