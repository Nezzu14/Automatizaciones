import pandas as pd

# Crear un DataFrame de ejemplo
data = {'Nombre': ['Juan', 'María', 'Pedro'],
        'Edad': [25, 30, 35],
        'Ciudad': ['Madrid', 'Barcelona', 'Sevilla']}
df = pd.DataFrame(data)

# Escribir el DataFrame en formato Parquet
df.to_parquet(r'C:\Users\migumart\OneDrive - Nokia\Archivos personales\Automatizacion Python\Templates_Python\Formato_Parquet.parquet')
