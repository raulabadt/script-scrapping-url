import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import messagebox

def scrape_and_save_to_excel():
    url = entry_url.get()

    # Realizar la solicitud HTTP
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.RequestException as e:
        messagebox.showerror("Error", f"Error al realizar la solicitud HTTP: {e}")
        return

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
        excel_filename = "datos_cotizaciones.xlsx"
        df.to_excel(excel_filename, index=False, engine='openpyxl')
        messagebox.showinfo("Éxito", f"Datos guardados en {excel_filename}")
    else:
        messagebox.showerror("Error", f"Error al realizar la solicitud HTTP. Código de estado: {response.status_code}")

# Crear la ventana principal
root = tk.Tk()
root.title("Scraping con GUI")

# Crear y posicionar los elementos en la ventana
label_url = tk.Label(root, text="URL de la página:")
label_url.grid(row=0, column=0, padx=10, pady=10)

entry_url = tk.Entry(root, width=50)
entry_url.grid(row=0, column=1, padx=10, pady=10)

button_scrape = tk.Button(root, text="Scrapear y Guardar en Excel", command=scrape_and_save_to_excel)
button_scrape.grid(row=1, column=0, columnspan=2, pady=10)

# Iniciar el bucle principal de la interfaz gráfica
root.mainloop()
