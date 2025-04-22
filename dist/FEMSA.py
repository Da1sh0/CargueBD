from sqlalchemy import create_engine
import pyodbc
import pandas as pd
from datetime import datetime
import os
import glob
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import ttk
import threading
import time
import webbrowser

# Configuración de conexión al servidor
SERVER = "DIIEGO-CAMIINO"
DATABASE = "FEMSA"
USERNAME = "Sa"
PASSWORD = "D13g0C4m11no*"

# Ruta donde se guardará el archivo
RUTA_EXPORTACION = r"../Reports"

# Crear la cadena de conexión compatible con SQLAlchemy
connection_string = f"mssql+pyodbc://{USERNAME}:{PASSWORD}@{SERVER}/{DATABASE}?driver=SQL+Server"

# Consulta SQL
query = """
SELECT CODIGO_SAP AS CODIGO_CLIENTE, 
IDENTIFICACION AS DOCUMENTO, 
NOMBRE_ESTABLECIMIENTO AS NOMBRE_NEGOCIO, 
RAZONSOCIAL AS NOMBRE_CLIENTE,  
DIRECCION, BARRIO_DESCRIPCION AS BARRIO, 
NOMBRE_MUNICIPIO AS CIUDAD, 
COALESCE (TELEFONO1, 'SIN REGISTRO') AS TELEFONO1, 
MAILCONTACTO AS CORREO_ELECTRONICO 
FROM FW_EMPRESAS (NOLOCK) 
INNER JOIN MUNICIPIO_CEDI ON FW_EMPRESAS.REGION_ID = MUNICIPIO_CEDI.ID 
WHERE ACTIVO = 1
"""
estado_proceso = ""


def actualizar_estado(mensaje):
    label_estado.config(text=mensaje)
    root.update_idletasks()


def generar_reporte():
    global exportando, start_time, estado_proceso
    exportando = True  # Inicia el temporizador
    start_time = time.time()

    try:
        # Crear un motor SQLAlchemy y conectar a la base de datos
        estado_proceso = "Conectando a la base de datos..."
        actualizar_tiempo()
        engine = create_engine(connection_string)
        with engine.connect() as conn:
            actualizar_estado("Consultando la base de datos...")
            df = pd.read_sql(query, conn)

        # Crear la carpeta si no existe
        if not os.path.exists(RUTA_EXPORTACION):
            os.makedirs(RUTA_EXPORTACION)

        # Guardar en Excel
        estado_proceso = "Guardando el archivo Excel..."
        actualizar_tiempo()
        fecha = datetime.now().strftime('%d_%m_%y')
        nombre_archivo = f"Base_Cliente {fecha}.xlsx"
        ruta_archivo = os.path.abspath(os.path.join(RUTA_EXPORTACION, nombre_archivo))  # Ruta final
        df.to_excel(ruta_archivo, index=False)
        wb = load_workbook(ruta_archivo)
        wb.save(ruta_archivo)  # Guardar en la ruta correcta

        estado_proceso = "Exportación completada."
        actualizar_tiempo()
        time.sleep(5)
    except Exception as e:
        actualizar_estado(f"Error: {e}")

    exportando = False  # Detiene el contador
    root.destroy()  # Cierra la ventana al finalizar


def actualizar_tiempo():
    # Actualiza el tiempo transcurrido en la etiqueta
    if exportando:
        elapsed_time = time.time() - start_time
        minutos = int(elapsed_time // 60)
        segundos = int(elapsed_time % 60)
        milisegundos = int((elapsed_time * 100) % 100)
        tiempo_str = f"{minutos:02}:{segundos:02}.{milisegundos:02}"

        label_estado.config(text=f"{estado_proceso}... {tiempo_str}")
        root.after(100, actualizar_tiempo)  # Llama a la función cada 100ms

def abrir_github(event):
    webbrowser.open_new("https://github.com/Da1sh0")

def mostrar_pantalla_carga():
    global root, label_estado, exportando, start_time

    root = tk.Tk()
    root.iconbitmap("femsa.ico")
    root.title("Femsa - Clientes")
    root.geometry("400x160")
    root.resizable(False, False)

    # Cargar la imagen
    logo = Image.open("Logo.png")  # Asegúrate de que el archivo esté en la misma carpeta
    logo = logo.resize((230, 50))  # Ajusta el tamaño si es necesario
    logo = ImageTk.PhotoImage(logo)

    # Mostrar la imagen en un Label
    label_logo = tk.Label(root, image=logo)
    label_logo.pack(pady=5)

    progreso = ttk.Progressbar(root, style="TProgressbar", orient="horizontal", length=300, mode="indeterminate")
    progreso.pack(pady=5)
    progreso.start()

    label_estado = tk.Label(root, text="Exportando la Base de Datos... 00:00.00", font=("Arial", 12))
    label_estado.pack(pady=3)

    label_version = tk.Label(root, text="v2.4 - By: Diiego Camiino", fg="gray", font=("Arial", 10, "italic"),cursor="hand2")
    label_version.pack(side="bottom", pady=5)
    label_version.bind("<Button-1>", abrir_github)

    exportando = True
    start_time = time.time()

    threading.Thread(target=generar_reporte, daemon=True).start()  # Inicia el proceso en otro hilo
    actualizar_tiempo()  # Inicia la actualización del tiempo

    root.mainloop()


if __name__ == "__main__":
    mostrar_pantalla_carga()

