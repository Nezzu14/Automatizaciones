import pandas as pd
from openpyxl import load_workbook
import os
from datetime import datetime

def xlookup():

    # ----Directorio de destino
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    xlsx_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xlsx"
    csv_file_path = directorio + "\\F&C GIC - SIG PC List - " + fecha + ".csv"

    # ----Se define el nombre y path del documento final
    doc_final_xlsx_path = directorio + "\\Prueba_Final" + fecha + ".xlsx"
    
    # ----Leer el archivo .csv
    df_csv = pd.read_csv(csv_file_path)
    
    # ----Leer el archivo .xlsx
    workbook = load_workbook(filename=xlsx_file_path)
    sheet = workbook.active
    df_xlsx = pd.read_excel(xlsx_file_path, engine="openpyxl", sheet_name=sheet.title)
    
    # ----Hacer el xlookup
    result = pd.merge(df_csv, df_xlsx, on="columna_comun")
    
    # ----Guardar el resultado en un nuevo archivo de Excel
    result.to_excel(doc_final_xlsx_path, index=False)



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#xlookup()
