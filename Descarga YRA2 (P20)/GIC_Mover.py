#TERCERO EN CORRER DE LOS ARCHIVOS GIC

import os
import shutil 

def Mover_GIC(Nombre_GIC):
        directorio_destino = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')
        descargas=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')
        ruta_archivo = descargas + "\\" + Nombre_GIC + ".slk"

        # Cortar el archivo al directorio de destino
        print(ruta_archivo)
        shutil.move(ruta_archivo, directorio_destino)