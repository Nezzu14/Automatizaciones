# Importing the Libraries
import win32com.client
import sys
import subprocess
import time
from datetime import datetime
import os
from tkinter import *
import disable
import win32gui
import GIC_Descarga

# ----This function will Login to SAP from the SAP Logon window
def saplogin(variante, username, password):

    print("==============================================================================================================")
    print("====INICIALIZACION DE -SAP LOGIN-")
    print("==============================================================================================================\n")

    try:

        # ----Path del ejecutable de SAP
        path = r"C:\Program Files (x86)\SAP\SAPGUI770\SAPgui\saplogon.exe"

        subprocess.Popen(path)
        hwnd = 0
        start_time = time.time()
        while not hwnd:
             hwnd = win32gui.FindWindow(None, 'SAP Logon 770')
             if time.time() - start_time > 30:
                return  # Si se supera el tiempo máximo de espera, se sale de la función
             time.sleep(0.5) 

        # ----Detecta la ventana de SAP
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        
        if application.Connections.Count==0 : 
            connection = application.OpenConnection("- P20 Production ERP Logistics and Finance", True)
            session = connection.Sessions(0)
            # ----Ingreso de Usuario y Contraseña
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            session.findById("wnd[0]").sendVKey(0)
        else: 
            if application.Connections.Count<6:
                  connection= application.Connections(0)
                  session = connection.Sessions(0)
                  session.CreateSession()
                  session=connection.Sessions(connection.Sessions.Count -1)
            else:
                print("Couldn't connect to application because sap reach the maximum number of sessions")

                print("==============================================================================================================")
                print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -TOPE DE SESIONES DE SAP ABIERTAS-")
                print("==============================================================================================================\n")
        
                print("========================================================================")
                print("----Couldn't connect to application because sap reach the maximum number of sessions")
                print("========================================================================\n")
        
                print(sys.exc_info())
                
                win= Tk()
        
                win.attributes('-topmost', True)
                #Set the geometry of frame
                win.geometry("400x140")
                win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
                win.title("TOPE DE SESIONES DE SAP ABIERTAS")
        
                def close_win():
                   win.destroy()
                
                #Create a text label
                Label(win,text="Couldn't connect to application because sap reach the maximum number of sessions", font=('Helvetica',10)).pack(pady=20)
        
                #Create Entry Widget for password
                
                #Create a button to close the window
                Button(win, text="Quit", font=('Helvetica bold',
                10),command=close_win).pack(pady=20, side="top")
                
                win.mainloop()
        
                print("==============================================================================================================")
                print("====FINALIZACION DE LA VENTANA EMERGENTE DE -TOPE DE SESIONES DE SAP ABIERTAS-")
                print("==============================================================================================================\n")

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
  
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
         
        username= username
        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA DESCARGA DEL REPORTE YRA2 Y DEL ARCHIVO GIC
        descargar(session, variante, username)
        #--------------------------------------------------------------------------------------------------------------------

         
        disable.Deshabiiltar_error()

    except:

        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
        print("==============================================================================================================\n")

        print("========================================================================")
        print("----Se ingresaron los datos de usuario y/o contraseña de forma incorrecta")
        print("========================================================================\n")

        print(sys.exc_info())
        
        win= Tk()

        win.attributes('-topmost', True)
        #Set the geometry of frame
        win.geometry("400x140")
        win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
        win.title("DATOS INCORRECTOS")

        def close_win():
           win.destroy()
        
        #Create a text label
        Label(win,text='Usuario y/o Contraseña incorrecta\nDarle a Quit dos veces y volver a iniciar el programa', font=('Helvetica',10)).pack(pady=20)

        #Create Entry Widget for password
        
        #Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=20, side="top")
        
        win.mainloop()

        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
        print("==============================================================================================================\n")

    finally:

        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -FIN DESCARGA YRA2-")
        print("==============================================================================================================\n")
        
        print("========================================================================")
        print("----Se termino la descarga del reporte YRA2 ejecutado en SAP")
        print("========================================================================\n")

        print(sys.exc_info())
        
        win= Tk()

        win.attributes('-topmost', True)
        #Set the geometry of frame
        win.geometry("400x120")
        win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
        win.title("FIN DESCARGA YRA2")

        def close_win():
           win.destroy()
        
        #Create a text label
        Label(win,text="Proceso Terminado", font=('Helvetica',10)).pack(pady=20)
        
        #Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=20, side="top")
        
        win.mainloop()

        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -FIN DESCARGA YRA2-")
        print("==============================================================================================================\n")

        session = None
        connection = None
        application = None
        SapGuiAuto = None
     
    print("==============================================================================================================")
    print("====FINALIZACION DE -SAP LOGIN-")
    print("==============================================================================================================\n")    
    
    exit()
       

def descargar(session, variante, username):
        
        print(username)
        username= username

        # Indicativo de la fecha actual
        fecha= "{:%Y_%m_%d}".format(datetime.now())

        print("======================================= ENTRANDO A YRA2 =======================================")
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"/YRA2"
        try:
           os.stat(directorio)
        except:
           os.mkdir(directorio)
        
        #Aca inicia el script the SAP hecho por SAP y se ejecuta entrando  a la transaccion de YRA2
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "yra2"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/usr/txtAENAME-LOW").text = variante
        session.findById("wnd[1]/usr/txtAENAME-LOW").setFocus()
        session.findById("wnd[1]/usr/txtAENAME-LOW").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
        print("======================================= ENTRANDO A YRA2 =======================================")
    
        #   Se empezara a descargar el Archivo de GIC
        GIC_Descarga.Descargar_GIC()

        # Path de como descargar el YRA2
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Pop up de ingreso de datos de la descarga
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directorio
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YRA2_TMOBILE_" + fecha + ".xls"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/mbar/menu[2]/menu[2]").select()
        session.findById("wnd[0]/mbar/menu[2]/menu[6]").select()


