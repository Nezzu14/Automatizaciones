import pandas as pd
import os
from datetime import datetime

def vlookup():

    print("==============================================================================================================")
    print("====INICIALIZACION DE -VLOOKUP-")
    print("==============================================================================================================\n")

    # ----Toma la fecha actual de hoy
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    xlsx_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\YRA2_TMOBILE_" + fecha + ".xlsx"
    csv_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"
    #csv_file_path = "C:\\Users\\migumart\\Desktop\\YRA2\\F&C GIC - SIG PC List - 2023_03_08.csv"

    # ----Se define el nombre y path del documento final
    doc_final_xlsx_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\Prueba_Final" + fecha + ".xlsx"

    print("========================================================================")
    print("Archivo xlsx = " + xlsx_file_path)
    print("Archivo csv = " + csv_file_path)
    print("Archivo final = " + doc_final_xlsx_path)
    print("========================================================================\n")

    print("========================================================================")
    print("----Inicio del vlookup entre reporte YRA2 y archivo GIC")
    print("========================================================================\n")

    # ----Define first DataFrame
    xlsx = pd.read_excel(xlsx_file_path)

    # ----Define second DataFrame
    csv = pd.read_csv(csv_file_path, encoding="latin")
    print(csv)

    vlookup_df = pd.merge(xlsx,
                         csv[['GIC', 'PC Business Group']],
                         on ='GIC',
                         how ='left')

    # ----View df1
    print(vlookup_df)

    # ----Save vlookup_df to Excel file
    vlookup_df.to_excel(doc_final_xlsx_path, index=False)

    print("========================================================================")
    print("----Fin del vlookup entre reporte YRA2 y archivo GIC")
    print("========================================================================\n")

    print("==============================================================================================================")
    print("====FINALIZACION DE -VLOOKUP-")
    print("==============================================================================================================\n")



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
vlookup()
