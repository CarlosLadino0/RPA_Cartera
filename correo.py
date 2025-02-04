from datetime import datetime, timedelta
from ejecutar_macro import ejecutar_macro
from utilidades import mostrar_mensaje
from threading import Thread
import tkinter as tk
import win32com.client
import os
import time
import pytz
import subprocess

def procesar_emails_y_guardar_informe():
    save_folder = "D:/Users/PC/Documents/CARTERA"
    os.makedirs(save_folder, exist_ok=True)

    try: 
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Outlook ya está abierto.")
    except Exception:
        print("Outlook no está abierto. Intentando abrirlo...")
        try:
            subprocess.Popen(["outlook.exe"], shell=True)
            time.sleep(5)
            outlook = win32com.client.Dispatch("Outlook.Application")
        except FileNotFoundError:
            print("No se encontró el ejecutable de Outlook. Asegúrate de que esté instalado correctamente.")
            mostrar_mensaje("Error", "No se pudo abrir Outlook. Verifica que esté instalado.", 2)
            return

    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)

    local_tz = pytz.timezone("America/Bogota")
    intentos = 0
    max_intentos = 10

    while intentos < max_intentos:  
        now = datetime.now(local_tz)
        print(f"Intento {intentos + 1} de {max_intentos} - Buscando correo... Fecha actual: {now}")

        try:
            inbox = namespace.GetDefaultFolder(6)  
            messages = inbox.Items.Restrict("[Unread] = True")  
            messages.Sort("[ReceivedTime]", True) 

            encontrado = False  

            for message in messages:
                received_time = message.ReceivedTime
                if not received_time.tzinfo:
                    received_time = pytz.utc.localize(received_time)
                received_time += timedelta(hours=5)

                sender_mail = message.SenderEmailAddress.lower()
                subject = message.Subject.strip()
                time_diff = now - received_time

                if sender_mail == "softseguros@softseguros.com" and subject == "Listado de Pagos por cobrar":
                    if time_diff <= timedelta(minutes=5):  
                        if message.Attachments.Count > 0:
                            for attachment in message.Attachments:
                                if attachment.FileName.endswith(".xlsx"):
                                    save_path = os.path.join(save_folder, "INFORME SS.xlsx")
                                    attachment.SaveAsFile(save_path)
                                    print(f"Archivo guardado en: {save_path}")
                                    message.Unread = False
                                    encontrado = True  
                                    break
                    if encontrado:
                        break  
            if encontrado:
                print("✅ Archivo descargado con éxito. Ejecutando macro...")
                mostrar_mensaje("Éxito", "El informe se ha descargado correctamente. Ejecutando macro...", 0)
                ejecutar_macro()
                return  

        except Exception as e:
            print(f"❌ Error procesando el correo: {e}")

        intentos += 1  
        if intentos < max_intentos:
            print("Correo no encontrado, actualizando bandeja y reintentando en 10 segundos...")
            time.sleep(10)  

    print("❌ No se encontró el correo después de 10 intentos. Cerrando el programa.")
    mostrar_mensaje("Error", "No se encontró el correo después de 10 intentos.", 2)