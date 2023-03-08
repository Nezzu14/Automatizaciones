import pandas as pd
import os
from datetime import datetime
import openpyxl

def vlookup():

    print("==============================================================================================================")
    print("====INICIALIZACION DE -VLOOKUP-")
    print("==============================================================================================================\n")

    # ----Directorio de destino
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    xlsx_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\YRA2_TMOBILE_" + fecha + ".xlsx"
    csv_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"

    # ----Se define el nombre y path del documento final
    doc_final_xlsx_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\Prueba_Final" + fecha + ".xlsx"
    
    # ----Cargar el archivo Excel
    workbook = openpyxl.load_workbook(xlsx_file_path)
    
    # ----Seleccionar la hoja que deseas modificar
    sheet = workbook["YRA2_TMOBILE_" + fecha]
    
    # ----Modificar el valor de una celda
    sheet['AK3'] = 'GIC'
    
    # ----Guardar los cambios en el archivo
    workbook.save(xlsx_file_path)

    print("========================================================================")
    print("----Nombre de la Celda _AK3_ se cambio de -GIC code- a -GIC-")
    print("========================================================================\n")

    print("========================================================================")
    print("----Inicio del vlookup entre reporte YRA2 y archivo GIC")
    print("========================================================================\n")

    # # ----leer los archivos de Excel y CSV en dos dataframes separados
    # df_excel = pd.read_excel(xlsx_file_path)
    # df_csv = pd.read_csv(csv_file_path)

    # # ----Realizar la búsqueda VLOOKUP y seleccionar solo la columna 4 del archivo CSV
    # df_resultado = pd.merge(df_excel[['GIC code']], df_csv[['GIC', 'columna4']], how='left')

    # # ----Exportar el dataframe resultado a un archivo CSV
    # df_resultado.to_csv(doc_final_xlsx_path, index=False)

    print("========================================================================")
    print("----Fin del vlookup entre reporte YRA2 y archivo GIC")
    print("========================================================================\n")

    print("==============================================================================================================")
    print("====FINALIZACION DE -VLOOKUP-")
    print("==============================================================================================================\n")



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
vlookup()
