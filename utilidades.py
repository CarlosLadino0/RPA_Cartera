from threading import Thread
import tkinter as tk
import ctypes

def mostrar_mensaje_temporal(titulo, mensaje, duracion=2000):
    def cerrar():
        root.destroy()
    root = tk.Tk()
    root.withdraw() 
    root.after(duracion, cerrar)  
    root.deiconify()
    root.title(titulo)
    label = tk.Label(root, text=mensaje, padx=20, pady=10)
    label.pack()
    root.mainloop()

def mostrar_mensaje(titulo, mensaje, tipo):
    Thread(target=mostrar_mensaje_temporal, args=(titulo, mensaje)).start()