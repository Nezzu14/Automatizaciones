#TERCERO EN CORRER DE LOS ARCHIVOS GIC

import os
import shutil 

def Mover_GIC():

        print("========================================================================================================================")
        print("====INICIALIZACION DEL MOVIMIENTO DEL ARCHIVO GIC DE DESCARGAS A YRA2")
        print("========================================================================================================================\n")

        # ----Obtenemos el nombre del archivo descargado automáticamente
        download_directory = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')  # ruta de tu carpeta de descargas
        filename_path = max([download_directory + "\\" + f for f in os.listdir(download_directory) if f.endswith('.slk')], key=os.path.getctime) #Nombre automatico del ultimo archivo descargado que sea .slk
        directorio_destino = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

        # ----Extraer el nombre del archivo de la ruta destino
        filename = os.path.basename(filename_path)

        print("========================================================================")
        print("----Download directory: " + download_directory)
        print("----Filename path: " + filename_path)
        print("========================================================================\n")
        
        # ----Check if file already exists
        if os.path.isdir(directorio_destino+'/'+ filename):
            print("========================================================================")
            print('----'+filename, '____ Exists in the destination path!')
            print("========================================================================\n")
            shutil.rmtree(directorio_destino+'/'+ filename)
        
        elif os.path.isfile(directorio_destino+'/'+ filename):
            os.remove(directorio_destino+'/'+ filename)
            print("========================================================================")
            print('----'+filename, '____ Deleted in', 'YRA2', 'becuase is duplicate')
            print("========================================================================\n")

        # ----Cortar el archivo al directorio de destino
        shutil.move(filename_path, directorio_destino)
        print("========================================================================")
        print("----Directorio destino: " + directorio_destino)
        print("========================================================================\n")

        print("========================================================================================================================")
        print("====FINALIZACION DEL MOVIMIENTO DEL ARCHIVO GIC DE DESCARGAS A YRA2")
        print("========================================================================================================================\n")

        return filename



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#Mover_GIC()