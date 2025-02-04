from utilidades import mostrar_mensaje
import win32com.client as win32
import time
import os

def ejecutar_macro():
    print("Ejecutando Macro")
    file_path = r'D:/Users/PC/Documents/CARTERA/PLANTILLA INFORME DE CARTERA.xlsm'
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        print("Excel abierto")

        workbook = excel.Workbooks.Open(file_path)
        print("Archivo abierto")

        excel.Application.Run('Módulo1.ActualizarPlantilla')
        print("Macro ejecutada")

        workbook.Save()
        print("Archivo guardado")
        
        workbook.Close(False)
        time.sleep(3)
        excel.Quit()
        print("Proceso completado")

        os.system("taskkill /f /im excel.exe")
        print("Excel cerrado completamente.")

        mostrar_mensaje("Éxito", "El informe se ha actualizado correctamente.", 0)

    except Exception as e:
        print(f"Error al ejecutar la macro: {e}")