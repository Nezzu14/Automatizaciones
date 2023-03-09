import pandas as pd
import os
from datetime import datetime
import shutil


def cambio_formato_csv():

    print("==============================================================================================================")
    print("====INICIALIZACION DE -CAMBIO FORMATO-")
    print("==============================================================================================================\n")

    # ----Indicativo la fecha actual de hoy
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    csv_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"
    xlsx_file_corregido_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".xlsx"
    filename_xlsx_corregido = "F&C GIC - SIG PC List - " + fecha + ".xlsx"

    print("========================================================================")
    print("Archivo csv = " + csv_file_path)
    print("========================================================================\n")

    print("========================================================================")
    print("----Modificacion de archivo .csv a .xlsx en proceso")
    print("========================================================================\n")

    # Leer el archivo CSV
    #csv = pd.read_csv(csv_file_path, encoding="cp1252")
    #csv = pd.read_csv(csv_file_path, encoding="ISO-8859-1")
    csv = pd.read_csv(csv_file_path, encoding="latin")
    #csv["GIC"] = csv["GIC"].astype(object)
    
    # ----Check if file already exists
    if os.path.isdir(xlsx_file_corregido_path):
        print("========================================================================")
        print('----'+filename_xlsx_corregido, '____ Exists in the destination path!')
        print("========================================================================\n")
        shutil.rmtree(xlsx_file_corregido_path)
    
    elif os.path.isfile(xlsx_file_corregido_path):
        os.remove(xlsx_file_corregido_path)
        print("========================================================================")
        print('----'+filename_xlsx_corregido, '____ Deleted in', 'YRA2', 'becuase is duplicate')
        print("========================================================================\n")

    # Escribir el archivo xlsx
    csv.to_excel(xlsx_file_corregido_path, index=False)

    print("========================================================================")
    print("----Modificacion terminada y archivo .xlsx guardado")
    print("----Archivo xlsx = " + xlsx_file_corregido_path)
    print("========================================================================\n")

    # ----Delete the original file
    os.remove(csv_file_path)

    print("========================================================================")
    print("----Archivo .xls antiguo eliminado")
    print("----Archivo xls = " + csv_file_path)
    print("========================================================================\n")

    print("==============================================================================================================")
    print("====FINALIZACION DE -CAMBIO FORMATO-")
    print("==============================================================================================================\n")



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#cambio_formato_csv()