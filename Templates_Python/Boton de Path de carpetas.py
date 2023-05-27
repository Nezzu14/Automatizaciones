import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import time
from tkinter import filedialog
from datetime import date
import downloads
from datetime import datetime
import os

seleccion = None


def ventana_finalizacion():
         # Crea una ventana emergente con un mensaje "Exitoso"
    root = ttk.Window()
    root.title("Programa exitoso")
    message = ttk.Label(root, text="Exitoso")
    message.pack(padx=20, pady=20)  
    button = ttk.Button(root, text="Cerrar", command=root.destroy)
    button.pack(pady=10)
    # Levanta la ventana por encima de todas las demás ventanas
    root.lift()
    root.attributes('-topmost', True)
 
    root.mainloop()


def capturar_datos():
    info_df=[]
    global seleccion
    group_list = entry_project.get("1.0", "end-1c")
    group_list = group_list.split("\n") # Dividir la cadena por saltos de línea
    group_list = [x for x in group_list if x != '']
    
 #   carpeta = entry_field_excel.get()

    fecha_inicio=date_start.entry.get()
    fecha_final = date_end.entry.get()
    
    periodo= period_spinbox.get()

    año = datetime.strptime(fecha_inicio, '%Y-%m-%d').year

    # Ruta del directorio de documentos
    ruta_documentos = os.path.expanduser("~/Documents")
    
    # Nombre de la carpeta a crear
    nombre_carpeta = "OBL_template"
    
    # Ruta completa de la carpeta a crear
    ruta_carpeta = os.path.join(ruta_documentos, nombre_carpeta)
    
    # Verificar si la carpeta ya existe antes de crearla
    if not os.path.exists(ruta_carpeta):
        # Crear la carpeta
        os.makedirs(ruta_carpeta)
        print("Carpeta creada con éxito.")
    else:
        print("La carpeta ya existe.")

    downloads.download_gd13(año,group_list, periodo,ruta_carpeta,fecha_inicio,fecha_final)

 #   downloads.download_yra2(fecha_inicio,fecha_final)

    print(año,group_list,fecha_inicio, fecha_final,periodo)



def seleccionar_carpeta():
    carpeta = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls .xlsm .csv")])
    if carpeta:
        # Si se seleccionó una carpeta, actualizar la entrada de texto
        entry_carpeta.delete(0, END)
        entry_carpeta.insert(0, carpeta)


# inicio de la aplicación usando el root window
root = ttk.Window()



# entradas de texto de la aplicación


label_project = ttk.Label(text="Group Keys")
label_project.pack(side="top", padx=20, pady=0)
entry_project = ttk.ScrolledText(width=30, height=5)
entry_project.pack(side="top", padx=20, pady=10)





label = ttk.Label(root, text="Fechas")
space1 = ttk.Label(root, text="")
space1.pack()

# label_excel = ttk.Label(root, text="Select an Excel file:")
# label_excel.pack()

# entry_field_excel = ttk.Entry(root)
# entry_field_excel.pack()

space1 = ttk.Label(root, text="")
space1.pack()

# browse_button = ttk.Button(root, text="Seleccionar drivers", command=browse_excel_file)
# browse_button.pack()


label = ttk.Label(root, text="Fechas")
space1 = ttk.Label(root, text="")
space1.pack()


# Get the current date and time
now = datetime.now()


dt2=date(now.year,now.month, now.day) # for startdate 

date_start  = ttk.DateEntry(dateformat='%Y-%m-%d',firstweekday=2,startdate=dt2)
date_start.pack()

space1 = ttk.Label(root, text="")
space1.pack()


date_end   = ttk.DateEntry(dateformat='%Y-%m-%d',firstweekday=2,startdate=dt2)
date_end .pack()


space1 = ttk.Label(root, text="")
space1.pack()

period_label = ttk.Label(root, text="Período")
period_label.pack()



period_spinbox = ttk.Spinbox(root, from_=1, to=12, width=12)
period_spinbox.pack()



# botones de la aplicación
b1 = ttk.Button(root, text="Ejecutar", bootstyle=SUNKEN, command=capturar_datos)
b1.pack(side="bottom", padx=5, pady=10)

folder_button = ttk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta)
folder_button.pack(side="top", padx=20, pady=10)

entry_carpeta = ttk.Entry(width=30)
entry_carpeta.pack(side="top", padx=20, pady=10)
label_carpeta = ttk.Label(text="Carpeta")
label_carpeta.pack(side="top", padx=20, pady=0)


root.pack_propagate(0)


root.geometry("700x800")  # aumenta la altura de la ventana
root.mainloop()

