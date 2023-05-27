import pandas as pd
from pathlib import Path
from datetime import datetime
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
import pyarrow as pa
import pyarrow.parquet as pq


fecha = "{:%Y_%m_%d}".format(datetime.now())



def capturar_datos():
    directorio = entry_carpeta.get()
    merge_obl(directorio)




info_df=[]




    

def merge_obl(folder_path):
        folder_path = Path(folder_path)
        xlsx_files = folder_path.glob('*.xlsx')
        dfs = []

        for file in xlsx_files:
            df = pd.read_excel(file,dtype=str)
            dfs.append(df)
    


        combined_df = pd.concat(dfs, ignore_index=True)


        combined_df.to_excel("ejemplo.xlsx", index=False)
        
        combined_df.to_csv("ejemplo.csv")
        combined_df.to_parquet("ejemplo.parquet")
        print("terminado")

        
        





def gd13(ruta_gd13, ruta_drivers):
    drivers= pd.read_excel(ruta_drivers,sheet_name="accounts")
    gd13= pd.read_excel(ruta_gd13)




def seleccionar_carpeta():
    carpeta = filedialog.askdirectory()
    if carpeta:
        # Si se seleccionó una carpeta, actualizar la entrada de texto
        entry_carpeta.delete(0, END)
        entry_carpeta.insert(0, carpeta)




root = ttk.Window()
folder_button = ttk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta)
folder_button.pack(side="top", padx=20, pady=10)

entry_carpeta = ttk.Entry(width=30)
entry_carpeta.pack(side="top", padx=20, pady=10)
label_carpeta = ttk.Label(text="Carpeta")
label_carpeta.pack(side="top", padx=20, pady=0)



b1 = ttk.Button(root, text="Ejecutar", bootstyle=SUNKEN, command=capturar_datos)
b1.pack(side="bottom", padx=5, pady=10)

      

root.title("T-MOBILE OBL")

root.geometry("200x300")  # aumenta la altura de la ventana
root.mainloop()


