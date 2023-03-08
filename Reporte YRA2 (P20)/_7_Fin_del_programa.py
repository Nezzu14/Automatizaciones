import sys
from datetime import datetime
from tkinter import *

        
def terminar_programa():
        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -TERMINAR PROGRAMA-")
        print("==============================================================================================================\n")
        
        print("========================================================================")
        print("----Se termino la automatizacion del Reporte YRA2 -FIN DEL PROGRAMA-")
        print("========================================================================\n")

        print(sys.exc_info())
        
        win= Tk()

        win.attributes('-topmost', True)
        # ----Set the geometry of frame
        win.geometry("400x70")
        win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
        win.title("REPORTE YRA2 - FIN DEL PROGRAMA")

        def close_win():
           win.destroy()
        
        # ----Create a text label
        Label(win,text="Proceso de Reporte YRA2 Terminado", font=('Helvetica',10,'bold')).pack(pady=5)
        
        # ----Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=5, side="top")
        
        win.mainloop()

        session = None
        connection = None
        application = None
        SapGuiAuto = None
    
        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -TERMINAR PROGRAMA-")
        print("==============================================================================================================\n")

        # ----Sale de ejecutar el PROGRAMA
        exit()