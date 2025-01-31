from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

url = "https://app.softseguros.com/"
driver = None
driver = webdriver.Chrome()

try: 
    driver.maximize_window()
    driver.set_page_load_timeout(20)
    driver.get(url)

    usuario = "claudiaarbelaez"
    campo_usuario = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "username"))
    )
    campo_usuario.send_keys(usuario)

    time.sleep(1)

    contrasena = "Correseguros*1"
    campo_contrasena = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.NAME, "password"))
    )
    campo_contrasena.send_keys(contrasena)
    campo_contrasena.send_keys(Keys.ENTER)

    time.sleep(1)
except Exception as e:
    print(f"Error al ingresar credenciales: {e}")

    '''
    Ellos agendaron una primera cita ayer pero como no pude estar, me dijeron que les enviara un correo hoy para re agendar la cita y hoy no contestaron
    '''