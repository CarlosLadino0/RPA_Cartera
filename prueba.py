'''
# URL del servidor
url = "https://app.softseguros.com/"

# Configuración y uso de Selenium
driver = None  # Inicializa el driver fuera del bloque try-except

try:
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.set_page_load_timeout(20)  # Tiempo máximo de espera: 10 segundos

    # Intenta acceder a la página
    try:
        driver.get(url)

        # Espera explícita para un elemento específico
        wait = WebDriverWait(driver, 10)
        campo = wait.until(
            EC.presence_of_element_located((By.NAME, "username"))  # Ajusta el selector según el campo real
        )
        print("La página cargó correctamente y el campo está disponible.")
    except TimeoutException:
        mostrar_error("La página tardó demasiado en cargar. Intenta nuevamente.")
        sys.exit()

    # Encuentra los campos de usuario y contraseña utilizando el atributo "name"
    username = driver.find_element(By.NAME, 'username')
    password = driver.find_element(By.NAME, 'password')

    # Ingresa las credenciales
    username.send_keys('claudiaarbelaez')  # Reemplaza con tu usuario
    password.send_keys('Correseguros*1')  # Reemplaza con tu contraseña

    # Simula presionar "Enter" para iniciar sesión
    password.send_keys(Keys.RETURN)

    time.sleep(5)

    print("Inicio de sesión completado.")
except WebDriverException as e:
    mostrar_error(f"Error al usar Selenium:{e}")
    '''










'''from datetime import datetime, timedelta
import win32com.client
import os
import time
import pytz
from actualizar import actualizar_bi
import subprocess
from pywinauto import Application

def open_outlook():
    """Abre Outlook con su ventana visible, forzando su inicio si no está activo."""
    try:
        # Verificar si Outlook ya está corriendo
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Outlook ya está abierto.")
    except Exception:
        # Si Outlook no está corriendo, intenta abrirlo
        print("Outlook no está abierto. Intentando abrirlo...")
        try:
            # Abrir Outlook explícitamente
            subprocess.Popen(["outlook.exe"], shell=True)
            time.sleep(5)  # Dar tiempo para que Outlook se inicie
            outlook = win32com.client.Dispatch("Outlook.Application")
        except FileNotFoundError:
            print("No se encontró el ejecutable de Outlook. Asegúrate de que esté instalado correctamente.")
            return None

    # Hacer visible la ventana de Outlook
    try:
        explorer = outlook.ActiveExplorer()
        if explorer is None:
            print("No hay ventana activa de Outlook. Abriendo nueva ventana...")
            inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)  # Carpeta Bandeja de Entrada
            inbox.Display()       
        else:
            print("Ventana de Outlook ya está activa.")
            
    except Exception as e:
        print(f"Error al mostrar la ventana de Outlook: {e}")
        return None

    return outlook


def check_outlook_emails(save_folder, processed_subjects, downloaded_files_set):
    # Conectar a Outlook
    #outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #inbox = outlook.Folders.Item("ingenierodesistemas@correseguros.co").Folders.Item("Bandeja de entrada")
    
    # Asegurarse de que Outlook esté abierto
    outlook = open_outlook()
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders.Item("ingenierodesistemas@correseguros.co").Folders.Item("Bandeja de entrada")
    
    # Buscar emails no leídos
    messages = inbox.Items
    messages = messages.Restrict("[Unread] = True")  # Filtrar solo correos no leídos

    # Ordenar por la fecha de recepción (más recientes primero)
    messages.Sort("[ReceivedTime]", True)
    
    downloaded_files = []

    # Obtener la fecha y hora actual en la zona horaria de Bogotá
    local_tz = pytz.timezone("America/Bogota")
    now = datetime.now(local_tz)

    print(f"Fecha actual: {now}")

    for message in messages:
        # Verificar si el correo fue recibido dentro de los últimos 30 minutos (ajustable)
        received_time = message.ReceivedTime

        # Asegurarse de que received_time tenga la zona horaria de UTC
        if not received_time.tzinfo:
            received_time = pytz.utc.localize(received_time)
            
        # Ajustar hora: Sumar 5 horas
        received_time += timedelta(hours=5)
        
        # Convertir la hora de recibo a la zona horaria local
        #received_time = received_time.astimezone(local_tz)

        # Depuración: Mostrar el remitente del correo
        sender_email = message.SenderEmailAddress
        #print(f"Asunto: {message.Subject} | Recibido: {received_time} | Remitente: {sender_email}")

        
        # Filtrar solo correos del remitente deseado
        if sender_email.lower() != "softseguros@softseguros.com":
            #print("Correo ignorado: Remitente no coincide.")
            continue
        
        # Calcular la diferencia de tiempo
        time_diff = now - received_time
        #print(f"Diferencia de tiempo: {time_diff}")

        # Procesar solo correos recibidos en los últimos 30 minutos
        if time_diff > timedelta(minutes=3):  # Ajusta el tiempo aquí si es necesario
            #print("Correo fuera del rango de tiempo. Ignorado.")
            continue

        if message.Subject in processed_subjects:
            continue  # Saltar correos ya procesados por su asunto

        if message.Attachments.Count > 0:  # Si tiene adjuntos
            for attachment in message.Attachments:
                if attachment.FileName.endswith(".xlsx"):  # Filtrar archivos Excel
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    sanitized_subject = "".join(c for c in message.Subject if c.isalnum() or c in (" ", "_")).strip()
                    new_file_name = f"{timestamp}{sanitized_subject}{attachment.FileName}"
                    save_path = os.path.join(save_folder, new_file_name)

                    # Evitar guardar duplicados si ya existe en downloaded_files_set
                    if save_path not in downloaded_files_set:
                        attachment.SaveAsFile(save_path)
                        downloaded_files.append((save_path, message.Subject))
                        downloaded_files_set.add(save_path)  # Agregar al conjunto para evitar duplicados
                        print(f"Archivo descargado: {save_path}")

            message.Unread = False  # Marcar correo como leído
            processed_subjects.add(message.Subject)  # Marcar correo como procesado

        # Parar el bucle si ya descargamos los dos archivos
        if len(downloaded_files) >= 2:
            break

    return downloaded_files

def main():
    save_folder = "C:\\Users\\JJ\\Documents\\CORREDORES DE SEGUROS\\INFORMES BI\\3. INFORME RENOVACIONES"
    os.makedirs(save_folder, exist_ok=True)

    print("Esperando correos con adjuntos...")

    downloaded_files = []
    processed_subjects = set()  # Para almacenar asuntos de correos ya procesados
    downloaded_files_set = set()  # Para almacenar rutas únicas de archivos descargados

    while len(downloaded_files) < 2:  # Esperar hasta que se descarguen 2 archivos
        new_files = check_outlook_emails(save_folder, processed_subjects, downloaded_files_set)
        downloaded_files.extend(new_files)
        
        # Si solo se ha encontrado un correo y no se han descargado 2 archivos, espera 60 segundos
        if len(downloaded_files) < 2:
            print("Esperando 10 segundos para comprobar más correos...")
            time.sleep(10)
    
    print("¡Descarga completada!")
    for file, subject in downloaded_files:
        print(f"Archivo guardado en: {file}, Asunto: {subject}")

    # Renombrar archivos según el asunto
    for file_path, subject in downloaded_files:
        if "Listado de Polizas y Anexos" in subject:
            new_name = "1. PRODUCCION Y ANEXOS.xlsx"
        elif "Listado de Polizas" in subject:
            new_name = "2. VENCIMIENTOS.xlsx"
        else:
            print(f"Asunto desconocido: {subject}. Archivo no renombrado.")
            continue

        new_path = os.path.join(save_folder, new_name)
        if os.path.exists(new_path):
            os.remove(new_path)  # Eliminar archivo existente
            print(f"Archivo eliminado: {new_path}")

        os.rename(file_path, new_path)
        print(f"Archivo renombrado a: {new_path}")

    actualizar_bi()

    
if _name_ == "_main_":
    main()'''