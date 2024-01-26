import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://www.interempresas.net/ovino/Articulos/538802-Lonja-de-Cereales-de-Barcelona-Cotizaciones-de-Cereal-(Semana-4-23-1-2024).html"

# Realizar la solicitud HTTP
response = requests.get(url)

# Verificar si la solicitud fue exitosa
if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')

    # Encontrar la tabla en la página (ajusta el selector según la estructura HTML)
    table = soup.find('table')

    # Extraer datos de la tabla y almacenarlos en una lista de listas
    data = []
    for row in table.find_all('tr'):
        cols = row.find_all(['td', 'th'])
        cols = [col.text.strip() for col in cols]
        data.append(cols)

    # Encontrar la longitud máxima de columnas en todas las filas
    max_columns = max(len(row) for row in data)

    # Rellenar las filas más cortas con valores nulos
    data = [row + [''] * (max_columns - len(row)) for row in data]

    # Crear un DataFrame de pandas con los datos
    df = pd.DataFrame(data[1:], columns=data[0])

    # Guardar el DataFrame en un archivo Excel
    df.to_excel("datos_cotizaciones.xlsx", index=False, engine='openpyxl')
    print("Datos guardados en datos_cotizaciones.xlsx")
else:
    print("Error al realizar la solicitud HTTP. Código de estado:", response.status_code)
