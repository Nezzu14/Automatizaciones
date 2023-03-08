import pandas as pd
import os
from datetime import datetime
import csv


def cambio_formato_csv():

    # ----Indicativo la fecha actual de hoy
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    csv_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"
    xlsx_file_corregido_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".xlsx"

    print("========================================================================")
    print("Archivo csv = " + csv_file_path)
    print("========================================================================\n")

    # Leer el archivo CSV
    df = pd.read_csv(csv_file_path, encoding="latin")

    # Escribir el archivo xlsx
    df.to_excel(xlsx_file_corregido_path, index=False)


    # ----Delete the original file
    os.remove(csv_file_path)



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#cambio_formato_csv()