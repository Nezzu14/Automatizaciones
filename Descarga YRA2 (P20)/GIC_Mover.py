#TERCERO EN CORRER DE LOS ARCHIVOS GIC

import os
import shutil 

def Mover_GIC(Path, Nombre_GIC):
        directorio_destino = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')
        ruta_archivo = Path + Nombre_GIC + ".xlsx"

        # Cortar el archivo al directorio de destino
        print(ruta_archivo)
        shutil.move(ruta_archivo, directorio_destino)