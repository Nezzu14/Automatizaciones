import win32com
import win32com.client
import os
from tkinter import filedialog
import pandas as pd
from pathlib import Path 
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

    
SapGuiAuto = win32com.client.GetObject('SAPGUI')     
if not type(SapGuiAuto) == win32com.client.CDispatch:
      pass
application = SapGuiAuto.GetScriptingEngine
if not type(application) == win32com.client.CDispatch:
    SapGuiAuto = None
    pass
connection = application.Children(0)
session    = connection.Children(0)


def new_folder(directorio):
   
   # Nombre de las carpetas a crear
   nombres_carpetas = ["Expanded", "Retracted"]
   
   # Crear las carpetas en el directorio especificado
   for nombre_carpeta in nombres_carpetas:
       ruta_carpeta = os.path.join(directorio, nombre_carpeta)
       os.makedirs(ruta_carpeta, exist_ok=True),

def merge_xslx_files(folder_path):
    folder_path = Path(folder_path)
    xlsx_files = folder_path.glob('*.xlsx')
    dfs = []
    for file in xlsx_files:
        df = pd.read_excel(file,dtype=str)
      
        
        dfs.append(df)
    
    combined_df = pd.concat(dfs, ignore_index=True)
    combined_df.to_excel(folder_path/Path("Merged_WBS.xlsx"),index=False)
    
    #classifier.clasificator(folder_path/Path("Merged_WBS.xlsx"))
    
    return combined_df


def convert_csv_files(folder_path,to_ignore):
    folder_path = Path(folder_path)
    csv_files = folder_path.glob('*.csv')
    for file in csv_files:
        df = pd.read_csv(file, sep='\t', usecols=range(3, 19),encoding="UTF-16",dtype=str,engine="python")
        df = df.iloc[:to_ignore]
        xlsx_path = folder_path / (file.stem + '.xlsx')
        print(df)
        df.to_excel(xlsx_path, index=False)


def unir_archivos(path_Retracted, path_Expanded, path_final):
    # Especifica los nombres de archivo de entrada y salida
    archivo1 = path_Retracted
    archivo2 = path_Expanded
    archivo_salida = path_final 

    # Carga los archivos de entrada en DataFrames de Pandas
    df1 = pd.read_excel(archivo1,dtype=str)
    df2 = pd.read_excel(archivo2,dtype=str)
    
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()
    # Aplicar la conversión de texto a números a todo el DataFrame
    # df1 = df1.applymap(lambda x: pd.to_numeric(x, errors='ignore') if isinstance(x, str) else x)
    # df2 = df2.applymap(lambda x: pd.to_numeric(x, errors='ignore') if isinstance(x, str) else x)
    
    with pd.ExcelWriter(archivo_salida) as writer:
       df1.to_excel(writer,sheet_name="Retracted",index=False)
       df2.to_excel(writer,sheet_name="Expanded",index=False)
    print(archivo_salida)
    

def download(wbs,directorio):


    wbs= str(wbs)
    print(wbs)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = "ca80"
    session.findById("wnd[0]/usr/ctxtP_POSID").text = wbs
    session.findById("wnd[0]/usr/ctxtP_DATE").setFocus()
    session.findById("wnd[0]/usr/ctxtP_DATE").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").select()
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = directorio+"\Expanded"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = wbs+"_expanded.csv"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[1]/btn[94]").press()
    session.findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").select()
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = directorio+"\Retracted"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = wbs+"_retracted.csv"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/mbar/menu[2]/menu[2]").select()
    session.findById("wnd[0]/mbar/menu[2]/menu[2]").select()




def capturar_datos():
    wbs_list = entry_project.get("1.0", "end-1c")
    wbs_list = wbs_list.split("\n") # Dividir la cadena por saltos de línea
    wbs_list = [x for x in wbs_list if x != '']
    
    directorio = entry_carpeta.get()
    new_folder(directorio)
    for wbs in wbs_list:
        download(wbs,directorio)

    convert_csv_files(directorio+"\Expanded",-1)
    convert_csv_files(directorio+"\Retracted",-4)
    
    merge_xslx_files(directorio+"\Expanded")
    merge_xslx_files(directorio+"\Retracted")

    unir_archivos(directorio+"\Retracted\Merged_WBS.xlsx",directorio+"\Expanded\Merged_WBS.xlsx",directorio+"/project_wip_report.xlsx")




def seleccionar_carpeta():
    carpeta = filedialog.askdirectory()
    if carpeta:
        # Si se seleccionó una carpeta, actualizar la entrada de texto
        entry_carpeta.delete(0, END)
        entry_carpeta.insert(0, carpeta)




# inicio de la aplicación usando el root window
root = ttk.Window()



# entradas de texto de la aplicación
entry_project = ttk.ScrolledText(width=30, height=5)

entry_project.pack(side="top", padx=20, pady=10)

label_project = ttk.Label(text="Project")

label_project.pack(side="top", padx=20, pady=0)



# botones de la aplicación
b1 = ttk.Button(root, text="Ejecutar", bootstyle=SUNKEN, command=capturar_datos)
b1.pack(side="bottom", padx=5, pady=10)

folder_button = ttk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta)
folder_button.pack(side="top", padx=20, pady=10)

entry_carpeta = ttk.Entry(width=30)
entry_carpeta.pack(side="top", padx=20, pady=10)
label_carpeta = ttk.Label(text="Carpeta")
label_carpeta.pack(side="top", padx=20, pady=0)



root.title("wbs wip")

root.geometry("700x800")  # aumenta la altura de la ventana
root.mainloop()


