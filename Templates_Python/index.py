import tkinter as tk
import ttkbootstrap as ttk
from PIL import ImageTk, Image
from ttkbootstrap.constants import *

botones_texto = ["OBL", "SULOS", "WIP", "CJ03", "GD13", "YRA2"]

root = ttk.Window(themename="solar", size=(900, 900))
root.title("GBS FPC Bogotá Oficce Automation toolkit")

frame = tk.Frame(root)
frame.pack(expand=True)

# Calcula la cantidad de botones en cada columna
cantidad_botones = len(botones_texto)
mitad = cantidad_botones // 2
if cantidad_botones % 2 == 1:
    mitad += 1

botones_izquierda = []
botones_derecha = []
for i, texto in enumerate(botones_texto):
    columna = i // mitad
    boton = ttk.Button(frame, width=20, text=texto)
    if columna == 0:
        botones_izquierda.append(boton)
    else:
        botones_derecha.append(boton)
    boton.grid(row=i % mitad, column=columna, pady=10)

image = Image.open(r"nokia.png")
img = image.resize((100, 30))
img = ImageTk.PhotoImage(img)
panel = ttk.Label(root, image=img)

# Ajustar la posición del widget de la etiqueta
panel.place(x=0, y=0)

root.mainloop()
