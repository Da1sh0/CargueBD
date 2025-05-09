from sqlalchemy import create_engine
import pandas as pd
from datetime import datetime
import sys
import os
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import ttk
import threading
import time
import webbrowser
import pyodbc

# Configuración de conexión al servidor
SERVER = "DIIEGO-CAMIINO"
DATABASE = "FEMSA"
USERNAME = "Sa"
PASSWORD = "D13g0C4m11no*"

# Ruta donde se guardará el archivo
RUTA_EXPORTACION_CLIENTES = r"./Reports/Clientes"
RUTA_EXPORTACION_EQUIPOS = r"./Reports/Repuestos"

# Crear la cadena de conexión compatible con SQLAlchemy
connection_string = f"mssql+pyodbc://{USERNAME}:{PASSWORD}@{SERVER}/{DATABASE}?driver=SQL+Server"

# Consulta SQL
queryC = """
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
queryE = """
SELECT  E.CODIGO_SAP , E.NUM_INVENTARIO, 
CE.DESCRIPCION AS ESTATUS_EQUIPO,
E.DESCRIPCION, E.SERIE,
CU.DESCRIPCION AS ESTATUS_USUARIO, 
E.GRUPO_PLANIFICADOR, E.CLIENTE_ID_DIRECTO, 
E.FECHA_ENTREGA, E.FECHA_INICIO_GARANTIA 
FROM FFA_EQUIPOS E (NOLOCK)
INNER JOIN FW_CATALOGOVALORES CE ON E.ESTATUS_SISTEMA = CE.CATVALID 
INNER JOIN FW_CATALOGOVALORES CU ON E.ESTATUS_USUARIO = CU.CATVALID
"""

estado_proceso = ""

def actualizar_estado(mensaje):
    global estado_proceso
    estado_proceso = mensaje
    def update():
        label_estado.config(text=mensaje)
    root.after(0, update)

def generar_reportes():
    global exportando, start_time, estado_proceso
    exportando = True
    start_time = time.time()

    try:
        # Crear motor
        actualizar_estado("Conectando a base de datos...")
        engine = create_engine(connection_string)

        # --- CLIENTES ---
        actualizar_estado("Consultando la base de Clientes...")
        with engine.connect() as conn:
            df_clientes = pd.read_sql(queryC, conn)
        actualizar_estado("Guardando archivo de Clientes...")
        if not os.path.exists(RUTA_EXPORTACION_CLIENTES):
            os.makedirs(RUTA_EXPORTACION_CLIENTES)
        fecha = datetime.now().strftime('%d_%m_%y')
        ruta_clientes = os.path.join(RUTA_EXPORTACION_CLIENTES, f"Base_Clientes {fecha}.xlsx")
        df_clientes.to_excel(ruta_clientes, index=False)
        actualizar_estado("Clientes exportados...")

        # --- EQUIPOS ---
        actualizar_estado("Consultando la base de Equipos...")
        with engine.connect() as conn:
            df_equipos = pd.read_sql(queryE, conn)
        actualizar_estado("Guardando archivo de Equipos...")
        if not os.path.exists(RUTA_EXPORTACION_EQUIPOS):
            os.makedirs(RUTA_EXPORTACION_EQUIPOS)
        ruta_equipos = os.path.join(RUTA_EXPORTACION_EQUIPOS, f"Base_Equipos {fecha}.xlsx")
        df_equipos.to_excel(ruta_equipos, index=False)
        actualizar_estado("Equipos exportados...")

        estado_proceso = "Exportación completada."
        actualizar_estado(estado_proceso)
        exportando = False
        root.after(5000, root.destroy)

    except Exception as e:
        exportando = False
        actualizar_estado(f"Error: {e}")

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

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Para cuando está empaquetado
    except Exception:
        base_path = os.path.abspath(".")  # Para cuando lo corres como .py
    return os.path.join(base_path, relative_path)

def mostrar_pantalla_carga():
    global root, label_estado, exportando, start_time

    root = tk.Tk()
    root.iconbitmap(resource_path("femsa.ico"))
    root.title("Femsa")
    root.geometry("400x160")
    root.resizable(False, False)

    # Cargar la imagen
    logo = Image.open(resource_path("Logo.png"))
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

    label_version = tk.Label(root, text="v3.5 - By: Diiego Camiino", fg="gray", font=("Arial", 10, "italic"),cursor="hand2")
    label_version.pack(side="bottom", pady=5)
    label_version.bind("<Button-1>", abrir_github)

    exportando = True
    start_time = time.time()

    threading.Thread(target=generar_reportes, daemon=True).start()
    actualizar_tiempo()  # Inicia la actualización del tiempo

    root.mainloop()

if __name__ == "__main__":
    mostrar_pantalla_carga()

