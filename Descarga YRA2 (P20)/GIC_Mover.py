#TERCERO EN CORRER DE LOS ARCHIVOS GIC

import os
import shutil 

def Mover_GIC():
        # Obtenemos el nombre del archivo descargado automáticamente
        download_directory = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')  # reemplaza con la ruta de tu carpeta de descargas
        filename = max([download_directory + "\\" + f for f in os.listdir(download_directory)], key=os.path.getctime)
        ruta_archivo = download_directory + filename + ".slk"
        directorio_destino = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

        print("=====================================download directory: " + download_directory)
        print("=====================================filename: " + filename)
        print("=====================================ruta archivo: " + ruta_archivo)
        print("=====================================directorio destino: " + directorio_destino)

        # Cortar el archivo al directorio de destino
        shutil.move(ruta_archivo, directorio_destino)
        print("=====================================ruta archivo pegado: " + ruta_archivo)