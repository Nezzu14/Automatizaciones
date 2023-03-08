import pandas as pd
import os
from datetime import datetime 


def cambio_formato_csv():

    # ----Toma la fecha actual de hoy
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    csv_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"

    # ----Se define el nombre y path del documento final
    doc_final_xlsx_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\Prueba_Final" + fecha + ".xlsx"

    print("========================================================================")
    print("Archivo csv = " + csv_file_path)
    print("Archivo final = " + doc_final_xlsx_path)
    print("========================================================================\n")

    # with open('archivo.csv', 'r') as file:
    #     # abrir archivo CSV en modo escritura
    #     with open('nuevo_archivo.csv', 'w', newline='') as new_file:
    #         # crear objeto csv.reader y csv.writer
    #         reader = csv.reader(file)
    #         writer = csv.writer(new_file, delimiter=',')
    #         # iterar sobre las filas del archivo
    #         for row in reader:
    #             # separar la fila por comas
    #             separated_row = row.split(',')
    #             # escribir la fila separada por comas en el nuevo archivo CSV
    #             writer.writerow(separated_row)

def cambio_formato_csv_a_xlsx():

    # ----Directorio de destino
    fecha= "{:%Y_%m_%d}".format(datetime.now())