#TERCERO EN CORRER DE LOS ARCHIVOS GIC

import os
import shutil 

def Mover_GIC():
        # Obtenemos el nombre del archivo descargado automáticamente
        download_directory = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')  # ruta de tu carpeta de descargas
        filename_path = max([download_directory + "\\" + f for f in os.listdir(download_directory) if f.endswith('.slk')], key=os.path.getctime) #Nombre automatico del ultimo archivo descargado que sea .slk
        directorio_destino = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

        print("========================================================================")
        print("=====================================download directory: " + download_directory)
        print("=====================================filename path: " + filename_path)
        print("========================================================================")
        
        # Cortar el archivo al directorio de destino
        shutil.move(filename_path, directorio_destino, overwrite=True) # Agregar el argumento overwrite=True para sobrescribir el archivo si existe
        print("========================================================================")
        print("=====================================directorio destino: " + directorio_destino)
        print("========================================================================")

#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#Mover_GIC()