from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from correo import procesar_emails_y_guardar_informe

def iniciar_sesion():
    url = "https://app.softseguros.com/"
    driver = None
    driver = webdriver.Chrome()

    try: 
        driver.maximize_window()
        driver.set_page_load_timeout(40)
        driver.get(url)

        # Usuario Softseguros (Cambiar según corresponda)
        usuario = ""
        campo_usuario = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "username"))
        )
        campo_usuario.send_keys(usuario)

        time.sleep(1)

        # Contraseña Softseguros (Cambiar según corresponda)
        contrasena = ""
        campo_contrasena = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "password"))
        )
        campo_contrasena.send_keys(contrasena)
        campo_contrasena.send_keys(Keys.ENTER)

        time.sleep(7)

    except Exception as e:
        print(f"Error al ingresar credenciales: {e}")

    try:

        btn_cobros = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="navigation"]/div[1]/div/nav/ul/li[5]/a'))
        )
        btn_cobros.click()
        driver.switch_to.active_element.send_keys(Keys.TAB)
        driver.switch_to.active_element.send_keys(Keys.ENTER)

        time.sleep(5)

    except Exception as e:
        print(f"Error al entrar al módulo de cobros: {e}")

    try: 
        btn_acciones_globales = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "dropdownMenu1"))
        )
        btn_acciones_globales.click()

        time.sleep(1)

        btn_exportar_excel = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a[data-id='cartera_por_cobrar']"))
        )
        btn_exportar_excel.click()

        time.sleep(1)

        campo_correo = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='ejemplo@email.com; ejemplo2@email.com']"))
        )
        # Correo técnico (Cambiar según corresponda)
        campo_correo.send_keys("")

        time.sleep(1)

        btn_ok = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//button[text()='OK']"))
        )
        btn_ok.click()

        time.sleep(5)

    except Exception as e:
        print(f"Error al exportar el Excel: {e}")

    finally:
        driver.quit()
        procesar_emails_y_guardar_informe()