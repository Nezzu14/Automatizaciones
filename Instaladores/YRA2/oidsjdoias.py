import tkinter as tk
from PIL import Image, ImageTk

class Reporte_YRA2:
    def __init__(self, master):
        master.attributes('-topmost', True)

        # ----Genera el titulo de la pantalla emergente
        self.master = master
        master.title('INICIO DESCARGA YRA2') #Titulo en el Pop up de ingresar Usuario y contraeña

        # ----Agregar la imagen a la ventana emergente
        image = Image.open("n-nokia.png")
        new_size = (100, 100)  # Especifica el nuevo tamaño deseado
        resized_image = image.resize(new_size)
        photo = ImageTk.PhotoImage(resized_image)    
        label = tk.Label(master, image=photo)
        label.pack()

        # ----Asignar la imagen al widget para evitar que se borre de la memoria
        label.image = photo

root = tk.Tk()
# ----Configura el color de fondo
root.configure(background='#FBFBFB')  # Blanco
# ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
root.iconbitmap(r'nokia.ico')

# ----Llamar la clase
Reporte_YRA2(master=root)
# ----Iniciar el bucle principal de la interfaz gráfica
root.mainloop()
