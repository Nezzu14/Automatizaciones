import tkinter as tk
import pickle
import os
import sys
import subprocess
import time
from datetime import datetime
from tkinter import *
import win32gui
import win32com.client as win32
import requests
import shutil
import pandas as pd
import win32con
import pywintypes
import win32com.client
import pyperclip


#---------------------------------------- Parte 1 ----------------------------------------
class InputForm:

    def on_entry_changed(self, input_str):
        
        # ----Esta def permite que solo lleguen a un maximo de 8 caracteres

        if len(input_str) > 10:
            self.radate_input.delete(10, 'end')
            self.createdon_input.delete(10, 'end')
            self.to_input.delete(10, 'end')

    def __init__(self, master):

        print("==============================================================================================================")
        print("====INICIALIZACION DE -REPORTE YRA2-")
        print("==============================================================================================================\n")

        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -DESCARGA YRA2-")
        print("==============================================================================================================\n")

        master.attributes('-topmost', True)

        # ----Genera el titulo de la pantalla emergente
        self.master = master
        master.title('INICIO DESCARGA YRA2') #Titulo en el Pop up de ingresar Usuario y contraeña

        # ----El directorio de los datos incriptados del usuario y password
        #directorio_login_bin = (r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\login_info.bin")
            # ----Si se quiere ejecutar en el computador  
        directorio_login_bin = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Nokia', 'Archivos personales', 'Automatizacion Python', 'Reporte YRA2 (P20)') + '\\login_info.bin'

        # ----Load saved username and password if they exist
        try:
            with open(directorio_login_bin, 'rb') as f:
                self.login_info = pickle.load(f)
        except:
            self.login_info = {'username': '', 'password': ''}

        # ----Create labels and input fields

        self.username_label = tk.Label(master, text='Fechas = dd.mm.yy', font=('Helvetica', 9, 'bold'))
        self.username_label.grid(row=0, column=1, padx=5, pady=5)
        
        self.username_label = tk.Label(master, text='Username:')
        self.username_label.grid(row=1, column=0, padx=5, pady=5)
        self.username_input = tk.Entry(master, width=16)
        self.username_input.insert(0, self.login_info['username'])
        self.username_input.grid(row=1, column=1, padx=5, pady=5)

        self.password_label = tk.Label(master, text='Password:')
        self.password_label.grid(row=2, column=0, padx=5, pady=5)
        self.password_input = tk.Entry(master, show='*', width=16)
        self.password_input.insert(0, self.login_info['password'])
        self.password_input.grid(row=2, column=1, padx=5, pady=5)

        self.radate_label = tk.Label(master, text='RA Date:')
        self.radate_label.grid(row=3, column=0, padx=5, pady=5)
        self.radate_input = tk.Entry(master, width=16)
        self.radate_input.grid(row=3, column=1, padx=5, pady=5)
        self.radate_input.bind('<KeyRelease>', lambda event: self.on_entry_changed(self.radate_input.get()))

        self.createdon_label = tk.Label(master, text='Created on:')
        self.createdon_label.grid(row=4, column=0, padx=5, pady=5)
        self.createdon_input = tk.Entry(master, width=16)
        self.createdon_input.grid(row=4, column=1, padx=5, pady=5)
        self.createdon_input.bind('<KeyRelease>', lambda event: self.on_entry_changed(self.createdon_input.get()))

        self.to_label = tk.Label(master, text='To:')
        self.to_label.grid(row=4, column=2, padx=5, pady=5)
        self.to_input = tk.Entry(master, width=16)
        self.to_input.grid(row=4, column=3, padx=5, pady=5)
        self.to_input.bind('<KeyRelease>', lambda event: self.on_entry_changed(self.to_input.get()))

        self.wbs_label = tk.Label(master, text='WBS\n(Separar por\n"Saltos de Linea"):')
        self.wbs_label.grid(row=5, column=0, padx=5, pady=5)
        self.wbs_input = tk.Text(master,height=5, width=12)
        self.wbs_input.grid(row=5, column=1, padx=5, pady=5)

        # ----Create submit button
        self.submit_button = tk.Button(master, text='Submit', command=self.submit)
        self.submit_button.grid(row=6, column=1, padx=5, pady=5)


    def submit(self):

        print("========================================================================")
        print("----Se presionó el boton SUBMIT ____ Procede a inicar -open_sap.saplogin(variante,username, password)-")
        print("========================================================================\n")

        # ----Get the values of the input fields and do something with them
        username = self.username_input.get()
        password = self.password_input.get()
        radate = self.radate_input.get()
        createdon = self.createdon_input.get()
        to = self.to_input.get()
        wbs = self.wbs_input.get('1.0', 'end-1c')
        
        # ----Esto es para volverlo lista, pero no se debe activar ya que el portapapleles no lee listas para ejecutar
        wbs_list = wbs.split("\n")or(",")or(" ")or(", ")or("  ")

        print("========================================================================")
        print("Username = " + username)
        print("Password = " + password)
        print("RA Date = " + radate)
        print("Created on = " + createdon)
        print("To = " + to)
        print("WBS's = ", wbs)
        print("WBS's = ", type(wbs))
        print("========================================================================\n")
        # ----Close the window and end the program pero si quieren seguir las varialbles se debe pner return al final del todo
        self.master.destroy()
        
        print("========================================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DESCARGA YRA2-")
        print("========================================================================================================================\n")

        # ----Save the login information
        self.login_info['username'] = username
        self.login_info['password'] = password
       
        # ----El directorio de los datos incriptados del usuario y password
        #directorio_login_bin = (r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\login_info.bin")
            # ----Si se quiere ejecutar en el computador        
        directorio_login_bin = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Nokia', 'Archivos personales', 'Automatizacion Python', 'Reporte YRA2 (P20)') + '\\login_info.bin'
        
        with open(directorio_login_bin, 'wb') as f:
            pickle.dump(self.login_info, f)

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA LA APERTURA DE SAP
        saplogin(username, password, radate, createdon, to, wbs_list) 
        #--------------------------------------------------------------------------------------------------------------------

        # ----Esto permite que se cierre la ventana emergente con self.master.destroy() pero que las variables y la ejecucion de los "def" no se borren
        return


#---------------------------------------- Parte 2 _ Open SAP ----------------------------------------
def saplogin(username, password, radate, createdon, to, wbs_list):

    try:
        # ----This function will Login to SAP from the SAP Logon window
        print("==============================================================================================================")
        print("====INICIALIZACION DE -SAP LOGIN-")
        print("==============================================================================================================\n")

        # ----Path del ejecutable de SAP
        path = r"C:\Program Files (x86)\SAP\SAPGUI770\SAPgui\saplogon.exe"

        # ----Esperar a que abra la pestaña de Log in de SAP para los ERP's 
        subprocess.Popen(path)
        hwnd = 0
        start_time = time.time()
        while not hwnd:
            hwnd = win32gui.FindWindow(None, 'SAP Logon 770')
            if time.time() - start_time > 30:
                return  # Si se supera el tiempo máximo de espera, se sale de la función
            time.sleep(0.5) 

        # ----Detecta la ventana de SAP
        SapGuiAuto = win32.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32.CDispatch:
            SapGuiAuto = None
            return

        # ----Si no hay sesiones abiertas se ejecuta SAP de cero
        if application.Connections.Count==0 : 
            connection = application.OpenConnection("- P20 Production ERP Logistics and Finance", True)
            session = connection.Sessions(0)

            # ----Ingreso de Usuario y Contraseña
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            session.findById("wnd[0]").sendVKey(0)

            # Obtener información de la sesión
            session_info = session.Info
            usuario=str(session_info.User)
            print("------------------------")
            print("----Lenght Usuario:",str(len(usuario)))
            print("------------------------")

            # Verificar si el inicio de sesión fue exitoso
            if len(usuario) == 0:
                    print("==============================================================================================================")
                    print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -REPORTE YRA2_DATOS INCORRECTOS-")
                    print("==============================================================================================================\n")

                    print("========================================================================")
                    print("----Se ingresaron los datos de usuario y/o contraseña de forma incorrecta")
                    print("========================================================================\n")

                    win= Tk()

                    win.attributes('-topmost', True)
                    # ----Set the geometry of frame
                    win.geometry("500x280")
                    #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
                    # ----Si se quiere ejecutar en el computador
                    win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
                    win.title("REPORTE YRA2 - DATOS INCORRECTOS")

                    def close_win():
                       win.destroy()

                    # ----Create a text label
                    Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
                    Label(win,text='1. Usuario, Contraseña y/o demas datos incorrectos', font=('Helvetica',10,'bold')).pack(pady=1)
                    Label(win,text='= Ejecute el programa de nuevo e ingrese los datos correctamente\n', font=('Helvetica',10)).pack(pady=0.1)
                    Label(win,text='2. Tiene seis sesiones abiertas, el cual es el maximo para SAP', font=('Helvetica',10,'bold')).pack(pady=1)
                    Label(win,text='= Cierre una de esas seis sesiones y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
                    Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
                    Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)

                    # ----Create a button to close the window
                    Button(win, text="Quit", font=('Helvetica bold',
                    10),command=close_win).pack(pady=10, side="top")

                    win.mainloop()

                    print("========================================================================")
                    print("----Se cerro la conexion de SAP")
                    # ----Cierra la pesteña de SAP ejecutada, y solo queda la de Log On
                    connection.CloseConnection()
                    print("----Se cerro la Pestaña de SAP Logon 770")
                    # Envía un mensaje WM_CLOSE a la ventana para cerrarla
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    print("========================================================================\n")

                    print("==============================================================================================================")
                    print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
                    print("==============================================================================================================\n")

                    # ----Sale de ejecutar el PROGRAMA PYTHON
                    exit()
            else:
                    print("========================================================================")
                    print("----" + usuario)
                    print("----Inicio de sesión exitoso.")
                    print("========================================================================")            

        else: 
            # ----Si ya hay sesiones abiertas con el acceso del usuario entonces se abrira una sesion aparte y se empezara a ejecitar el proceso de descarga de YRA2
            # ----Se abrira solo hasta el maximo de 6 sesiones, si ya hay 6 sesiones abiertas entonces arrojara un error, el cual es el de "except"
            if application.Connections.Count<2:
                connection= application.Connections(0)
                session = connection.Sessions(0)
                session.CreateSession()
                #session=connection.Sessions(connection.Sessions.Count -1)
                print("========================================================================")
                print("----Ususario ya activo - Se abrio  sesion")
                print("========================================================================\n")
            else:
                print("==============================================================================================================")
                print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -MAS DE DOS VENTAS DE INICIO DE SESION ABIERTAS-")
                print("==============================================================================================================\n")

                print("========================================================================")
                print("----Hay dos o mas conexiones abiertas (Inicios de sesion)")
                print(sys.exc_info())
                print("========================================================================\n")

                win= Tk()

                win.attributes('-topmost', True)
                # ----Set the geometry of frame
                win.geometry("540x240")
                #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
                # ----Si se quiere ejecutar en el computador
                win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
                win.title("REPORTE YRA2 - MAS DE DOS VENTAS DE INICIO DE SESION ABIERTAS")

                def close_win():
                   win.destroy()

                # ----Create a text label
                Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
                Label(win,text='1. Hay dos o mas conexiones abiertas (Inicios de sesion)', font=('Helvetica',10,'bold')).pack(pady=1)
                Label(win,text='= Cierre las ventanas y ejecute el programa de nuevo\n', font=('Helvetica',10)).pack(pady=0.1)
                Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
                Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)

                # ----Create a button to close the window
                Button(win, text="Quit", font=('Helvetica bold',
                10),command=close_win).pack(pady=1, side="top")

                win.mainloop()

                print("==============================================================================================================")
                print("====FINALIZACION DE LA VENTANA EMERGENTE DE -MAS DE DOS VENTAS DE INICIO DE SESION ABIERTAS-")
                print("==============================================================================================================\n")

                # ----Sale de ejecutar el PROGRAMA
                exit()                



        if not type(connection) == win32.CDispatch:
            application = None
            SapGuiAuto = None
            return
    
        if not type(session) == win32.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        print("==============================================================================================================")
        print("====FINALIZACION DE -SAP LOGIN-")
        print("==============================================================================================================\n")  
    except pywintypes.com_error as e_sap:
        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -REPORTE YRA2_DATOS INCORRECTOS-")
        print("==============================================================================================================\n")

        print("========================================================================")
        print("----Error de SAP: ")
        #Este print es para que me muestre el Error en especifico si no lo se: print(sys.exc_info()), ya que de esta forma se sabe que el error generado es "pywintypes.com_error"
        print(str(e_sap))
        print("========================================================================\n")

        win= Tk()

        win.attributes('-topmost', True)
        # ----Set the geometry of frame
        win.geometry("500x280")
        #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
        # ----Si se quiere ejecutar en el computador
        win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
        win.title("REPORTE YRA2 - DATOS INCORRECTOS")

        def close_win():
            win.destroy()

        # ----Create a text label
        Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
        Label(win,text='1. Error de SAP', font=('Helvetica',10,'bold')).pack(pady=1)
        Label(win,text='= Una vez cerrada la ventana de error de SAP y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
        Label(win,text='2. No se encuentra conectado a la VPN de CISCO o la red de Nokia', font=('Helvetica',10,'bold')).pack(pady=1)
        Label(win,text='= Conectese a la VPN de CISCO o la red de Nokia y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
        Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
        Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)

        # ----Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=10, side="top")

        win.mainloop()

        print("========================================================================")
        print("----Se cerro la Pestaña de SAP Logon 770")
        # Envía un mensaje WM_CLOSE a la ventana para cerrarla
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        print("========================================================================\n")

        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
        print("==============================================================================================================\n")

        # ----Sale de ejecutar el PROGRAMA PYTHON
        exit()

    #--------------------------------------------------------------------------------------------------------------------
    # <<<<<<<<<SE EJECUTA DESCARGA DEL REPORTE YRA2 Y DEL ARCHIVO GIC
    Path_YRA2_SAP(session, username, radate, createdon, to, wbs_list, connection, hwnd)
    #--------------------------------------------------------------------------------------------------------------------

    session = None
    connection = None
    application = None
    SapGuiAuto = None

    #==============================================================================================================
    #====FINALIZACION DE -SAP LOGIN- \\\\CODIGO
    #==============================================================================================================    

def Path_YRA2_SAP(session, username, radate, createdon, to, wbs_list, connection, hwnd):
        
        print("==============================================================================================================")
        print("====INICIALIZACION DE -PATH YRA2 SAP-")
        print("==============================================================================================================\n")  

        print("========================================================================")
        print("VERIFICACION DE DATOS")
        print("Username:",username)
        print("Fecha RA Date:", radate)
        print("Fecha Created On:", createdon)
        print("Fecha To:", to)
        print("========================================================================\n")

        # ----Indicativo de la fecha actual
        fecha= "{:%Y_%m_%d}".format(datetime.now())

        print("========================================================================")
        print("----Entrando a YRA2 en SAP")
        print("========================================================================\n")

        # ------------------------------------ Inicio Try
        # ----Check if file already exists
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"/YRA2"
        try:
           os.stat(directorio)
        except:
           os.mkdir(directorio)
        # ------------------------------------ Fin Try


        # ------------------------------------ Inicio Try
        # ----Aca inicia el script the SAP hecho por SAP y se ejecuta entrando  a la transaccion de YRA2
        try:
            # ----Se ejecuta antes en SAP logon para que pueda dar error
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "yra2"
            session.findById("wnd[0]").sendVKey (0)

            session.findById("wnd[0]/usr/ctxtS_PSPID-LOW").text = "G-0609A3EZ1" # ----Top WBS por defecto para que deje abrir la ventana emergente y pegar todas las WBS
            session.findById("wnd[0]/usr/ctxtP_DATUM").text = radate # ----RA Date
            session.findById("wnd[0]/usr/ctxtP_DATUM").caretPosition = 10
            session.findById("wnd[0]/usr/btnB_EXPDOC").press() # ----Presiona el boton para abrir las opciones de las fechas
            session.findById("wnd[0]/usr/ctxtS_CPUDT-LOW").text = createdon # ----Created on
            session.findById("wnd[0]/usr/ctxtS_CPUDT-HIGH").text = to # ----To
            session.findById("wnd[0]/usr/ctxtS_CPUDT-HIGH").setFocus()
            session.findById("wnd[0]/usr/ctxtS_CPUDT-HIGH").caretPosition = 10

            # ----Esto hace que en el portapapeles se peguen los WBS y puedan ser ingresados en la funcion de SAP de pegar el portapapeles
            #pyperclip.copy(wbs)
            print("========================================================================")
            print("VERIFICACION DE DATOS")
            print("WBS List: ", wbs_list)
            print("========================================================================")

            session.findById("wnd[0]/usr/btn%_S_PSPID_%_APP_%-VALU_PUSH").press() # ----Abre la ventana para que pongan varios WBS
            for i, element in enumerate(wbs_list):
                session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," + str(i) + "]").text = str(element)
                print("========================================================================")
                print("Linea en la columna y WBS de linea en la columna: ", i, element)
                print("========================================================================")
            session.findById("wnd[1]/tbar[0]/btn[8]").press() # ----Completa la pestaña de varias WBS
            
            session.findById("wnd[0]/tbar[1]/btn[8]").press() # ----Ejecuta la carga del reporte YRA2

            print("========================================================================")
            print("----Carga del reporte YRA2 en SAP")
            print("========================================================================\n")

            print("========================================================================")
            print("----Inicia el proceso de descarga del reporte YRA2")
            print("========================================================================\n")

            # ----Path de como descargar el YRA2
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # ----Pop up de ingreso de datos de la descarga
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = directorio
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YRA2_TMOBILE_" + fecha + ".xls"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            session.findById("wnd[0]/mbar/menu[2]/menu[2]").select()
            session.findById("wnd[0]/mbar/menu[2]/menu[6]").select()
        except:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -DATOS INGRESADOS INCORRECTOS EN YRA2-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----Se ingresaron los datos de forma incorrecta de la transaccion YRA2")
            print(sys.exc_info())
            print("========================================================================\n")

            win= Tk()

            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("540x280")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - DATOS INCORRECTOS")

            def close_win():
               win.destroy()

            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. Fechas incorrectas o formato de fechas incorrectas', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Ejecute el programa de nuevo e ingrese las fechas correctamente\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='2. WBS mal ingresadas o WBS inexistente', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Revise las WBS, luego vuelva a ejecutar el programa e ingrese las WBS correctamente\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
    
            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=1, side="top")

            win.mainloop()

            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE --DATOS INGRESADOS INCORRECTOS EN YRA2--")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----Se cerro la conexion de SAP")
            # ----Cierra la pesteña de SAP ejecutada, y solo queda la de Log On
            connection.CloseConnection()
            print("----Se cerro la Pestaña de SAP Logon 770")
            # Envía un mensaje WM_CLOSE a la ventana para cerrarla
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            print("========================================================================\n")

           
            # ----Sale de ejecutar el PROGRAMA
            exit()
        # ------------------------------------ Fin Try

        print("==============================================================================================================")
        print("====FINALIZACION DE -PATH YRA2 SAP-")
        print("==============================================================================================================\n")

        #==============================================================================================================
        #====FINALIZACION DE -PATH YRA2 SAP- \\\\CODIGO
        #==============================================================================================================  


        #====================================================================================================================
        # ----INICIO DEL RESTO DEL PROCESO DE DOCUMENTOS DEL REPORTE YRA2----
        
        #--------------------------------------------------------------------------------------------------------------------
        # ---Deshabiiltar_error()
        try:
            print("========================================================================")
            print("----Una vez guardado el reporte YRA2 en .xls, se corregira en .xlsx con -Correccion_formato._3_Deshabiiltar_error()-")
            print("========================================================================\n")

            #--------------------------------------------------------------------------------------------------------------------
            # <<<<<<<<<SE CORREGIRA EL FORMATO DEL REPORTE YRA2
            Deshabiiltar_error()
            #--------------------------------------------------------------------------------------------------------------------
        except:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -ERROR EN DESHABILITAR ERROR-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----No se encontro el documento de YRA2")
            print(sys.exc_info())
            print("========================================================================\n")

            win= Tk()

            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("540x240")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - NO SE ENCONTRO EL DOC DE YRA2")

            def close_win():
               win.destroy()

            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. No se encontro el documento de YRA2', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Ejecute el programa de nuevo\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
    
            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=1, side="top")

            win.mainloop()

            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE --ERROR EN DESHABILITAR ERROR--")
            print("==============================================================================================================\n")
           
            # ----Sale de ejecutar el PROGRAMA
            exit()
        #--------------------------------------------------------------------------------------------------------------------

        #--------------------------------------------------------------------------------------------------------------------
        # ---Descargar_GIC()
        try:
            print("========================================================================")
            print("----Una vez guardado correctamente el reporte YRA2, se empieza a ejecutar como segundo proceso -_4_GIC_Descarga.Descargar GIC-")
            print("========================================================================\n")

            #--------------------------------------------------------------------------------------------------------------------
            # <<<<<<<<<SE EMPEZARA A DESCARGAR EL ARCHIVO GIC
            Descargar_GIC()
            #--------------------------------------------------------------------------------------------------------------------
        except:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -ERROR EN DESCARGA GIC-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----No se pudo descargar el archivo GIC")
            print(sys.exc_info())
            print("========================================================================\n")

            win= Tk()

            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("540x240")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - ERROR EN DESCARGA GIC")

            def close_win():
               win.destroy()

            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. No se pudo descargar el archivo GIC', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Ejecute el programa de nuevo\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
    
            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=1, side="top")

            win.mainloop()

            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE --ERROR EN DESCARGA GIC--")
            print("==============================================================================================================\n")
           
            # ----Sale de ejecutar el PROGRAMA
            exit()
        #--------------------------------------------------------------------------------------------------------------------

        #--------------------------------------------------------------------------------------------------------------------
        # ----cambio_formato_csv()
        try:
            print("========================================================================")
            print("----Una vez descargado el archivo GIC, se empieza a corregir el formato del archivo .csv -_5_GIC_Descarga.Descargar GIC-")
            print("========================================================================\n")

            #--------------------------------------------------------------------------------------------------------------------
            # <<<<<<<<<SE EMPEZARA A CORREGIR EL ARCHIVO GIC
            cambio_formato_csv()
            #--------------------------------------------------------------------------------------------------------------------
        except:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -ERROR EN CAMBIO FORMATO GIC-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----No se encontro el documento del GIC")
            print(sys.exc_info())
            print("========================================================================\n")

            win= Tk()

            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("540x240")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - ERROR EN CAMBIO FORMATO DE GIC")

            def close_win():
               win.destroy()

            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. No se encontro el documento del GIC', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Ejecute el programa de nuevo\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
    
            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=1, side="top")

            win.mainloop()

            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE --ERROR EN CAMBIO FORMATO GIC--")
            print("==============================================================================================================\n")
           
            # ----Sale de ejecutar el PROGRAMA
            exit()
        #--------------------------------------------------------------------------------------------------------------------

        #--------------------------------------------------------------------------------------------------------------------
        # ----vlookup()
        try:
            print("========================================================================")
            print("----Una vez guardado correctamente el reporte YRA2 y el archivo GIC, se empieza a ejecutar el Vlookup entre ambos archivos -_6_Clasificador.vlookup-")
            print("========================================================================\n")

            #--------------------------------------------------------------------------------------------------------------------
            # <<<<<<<<<SE EJECUTARA EL VLOOKUP ENTRE EL ARCHIVO .CSV A .XLSX
            vlookup()
            #-------------------------------------------------------------------------------------------------------------------- 

            print("==============================================================================================================")
            print("====FINALIZACION DE -REPORTE YRA2-")
            print("==============================================================================================================\n")
        except:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -ERROR EN VLOOKUP-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----No se pudo hacer el vlookup entre el archivo GIC y el Reporte YRA2")
            print(sys.exc_info())
            print("========================================================================\n")

            win= Tk()

            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("540x240")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - ERROR EN DESCARGA GIC")

            def close_win():
               win.destroy()

            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. No se pudo hacer el vlookup entre el archivo GIC y el Reporte YRA2', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Ejecute el programa de nuevo\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
    
            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=1, side="top")

            win.mainloop()

            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE --ERROR EN VLOOKUP--")
            print("==============================================================================================================\n")
           
            # ----Sale de ejecutar el PROGRAMA
            exit()
        #--------------------------------------------------------------------------------------------------------------------

        #--------------------------------------------------------------------------------------------------------------------
        # ----terminar_programa()
        try:
            print("========================================================================")
            print("----Una vez finalizado el proceso del Reporte YRA2 se terminara el programa -_7_Fin_del_programa.terminar_programa-")
            print("========================================================================\n")

            #--------------------------------------------------------------------------------------------------------------------
            # <<<<<<<<<SE EJECUTARA LA VENTANA DE FINALIZACION DE PROGRAMA
            terminar_programa(connection, hwnd)
            #--------------------------------------------------------------------------------------------------------------------
        except:
            print("==============================================================================================================")
            print("xd no se que pudo haber pasado en este punto")
            print("==============================================================================================================")
        #--------------------------------------------------------------------------------------------------------------------
        
        # ----FIN DEL RESTO DEL PROCESO DE DOCUMENTOS DEL REPORTE YRA2----
        #====================================================================================================================


#---------------------------------------- Parte 3 _ Correccion Formato .XLSX ----------------------------------------
def Deshabiiltar_error():
        
        print("==============================================================================================================")
        print("====INICIALIZACION DE -DESHABILITAR ERROR-")
        print("==============================================================================================================\n")  

        # ----Toma la fecha actual de hoy
        fecha= "{:%Y_%m_%d}".format(datetime.now())

        # ----Directorio de destino
        print(fecha)
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

        # ----Se definen los paths de los archivos, el archivo original y el archivo al que se quiere convertir
        original_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xls"
        modified_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xlsx"

        print("========================================================================")
        print("----Modificacion de archivo .xls a .xlsx en proceso")
        print("========================================================================\n")

        # ----Create an instance of the Excel application object and open the original file
        excel = win32.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(original_file_path)
        excel.DisplayAlerts = False

        # ----Save the modified workbook as an XLSX file
        wb.SaveAs(modified_file_path, FileFormat=51) 

        # ----Close the workbook
        wb.Close()

        # ----Exit the Excel application
        #excel.Quit()

        print("========================================================================")
        print("----Modificacion terminada y archivo .xlsx guardado")
        print("Archivo xlsx = " + modified_file_path)
        print("========================================================================\n")

        # ----Delete the original file
        os.remove(original_file_path)

        print("========================================================================")
        print("----Archivo .xls antiguo eliminado")
        print("Archivo xls = " + original_file_path)
        print("========================================================================\n")

        print("==============================================================================================================")
        print("====FINALIZACION DE -DESHABILITAR ERROR-")
        print("==============================================================================================================\n") 

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA LA CORRECCION DENTRO DE LOS ARCHIVOS DE REPORTE YRA2 Y DEL ARCHIVO GIC
        #correciones_lineas_xlsx()
        #--------------------------------------------------------------------------------------------------------------------

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA LA CORRECCION DENTRO DE LOS ARCHIVOS DE REPORTE YRA2 Y DEL ARCHIVO GIC
        #correciones_lineas_win32()
        #--------------------------------------------------------------------------------------------------------------------


#---------------------------------------- Parte 4 _ GIC Descarga ----------------------------------------
def Descargar_GIC(): 

    print("==============================================================================================================")
    print("====INICIALIZACION DE LA DESCARGA DEL ARCHIVO GIC")
    print("==============================================================================================================\n")

    print("========================================================================")
    print("----Descargando Archivo GIC")
    print("========================================================================\n")

    # ----Define la variable url que contiene la dirección URL del archivo que se va a descargar
    url = "http://dataq-prod.int.net.nokia.com:7780/pls/apex/f?p=115:8::CSV::::"    

    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Define la variable filename que contiene el nombre que se le dará al archivo descargado en el sistema local.
    filename_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"
    filename = "F&C GIC - SIG PC List_" + fecha + ".csv"

    # ----Te avisa si ya descargo el archivo
    response = requests.get(url)

    print("========================================================================")
    print("----Archivo GIC descargado en la Carpeta YRA2 del Escritorio")
    print("----"+filename)
    print("========================================================================\n")

    # ----Verifica si la carpeta "mi_carpeta" existe en el sistema de archivos utilizando la función "os.path.exists()". Si la carpeta no existe, se crea utilizando la función "os.makedirs()".
    directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"\\YRA2"
    try:
       os.stat(directorio)
    except:
       os.mkdir(directorio)

    # ----Check if file already exists
    if os.path.isdir(filename_path):
        print("========================================================================")
        print('----'+filename, '____ Exists in the destination path!')
        print("========================================================================\n")
        shutil.rmtree(filename_path)
    
    elif os.path.isfile(filename_path):
        os.remove(filename_path)
        print("========================================================================")
        print('----'+filename, '____ Deleted in', 'YRA2', 'becuase is duplicate')
        print("========================================================================\n")

    with open(filename_path, "wb") as f:
        f.write(response.content)

    print("==============================================================================================================")
    print("====FINALIZACION DE LA DESCARGA DEL ARCHIVO GIC")
    print("==============================================================================================================\n")


#---------------------------------------- Parte 5 _ Correccion Formato .CSV ----------------------------------------
def cambio_formato_csv():

    print("==============================================================================================================")
    print("====INICIALIZACION DE -CAMBIO FORMATO-")
    print("==============================================================================================================\n")

    # ----Indicativo la fecha actual de hoy
    fecha= "{:%Y_%m_%d}".format(datetime.now())
    
    # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
    csv_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".csv"
    xlsx_file_corregido_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".xlsx"
    filename_xlsx_corregido = "F&C GIC - SIG PC List - " + fecha + ".xlsx"

    print("========================================================================")
    print("Archivo csv = " + csv_file_path)
    print("========================================================================\n")

    print("========================================================================")
    print("----Modificacion de archivo .csv a .xlsx en proceso")
    print("========================================================================\n")

    # Leer el archivo CSV
    #csv = pd.read_csv(csv_file_path, encoding="cp1252")
    #csv = pd.read_csv(csv_file_path, encoding="ISO-8859-1")
    csv = pd.read_csv(csv_file_path, encoding="latin")
    #csv["GIC"] = csv["GIC"].astype(object)
    
    # ----Check if file already exists
    if os.path.isdir(xlsx_file_corregido_path):
        print("========================================================================")
        print('----'+filename_xlsx_corregido, '____ Exists in the destination path!')
        print("========================================================================\n")
        shutil.rmtree(xlsx_file_corregido_path)
    
    elif os.path.isfile(xlsx_file_corregido_path):
        os.remove(xlsx_file_corregido_path)
        print("========================================================================")
        print('----'+filename_xlsx_corregido, '____ Deleted in', 'YRA2', 'becuase is duplicate')
        print("========================================================================\n")

    # Escribir el archivo xlsx
    csv.to_excel(xlsx_file_corregido_path, index=False)

    print("========================================================================")
    print("----Modificacion terminada y archivo .xlsx guardado")
    print("----Archivo xlsx = " + xlsx_file_corregido_path)
    print("========================================================================\n")

    # ----Delete the original file
    os.remove(csv_file_path)

    print("========================================================================")
    print("----Archivo .xls antiguo eliminado")
    print("----Archivo xls = " + csv_file_path)
    print("========================================================================\n")

    print("==============================================================================================================")
    print("====FINALIZACION DE -CAMBIO FORMATO-")
    print("==============================================================================================================\n")


#---------------------------------------- Parte 6 _ Clasificador ----------------------------------------
def vlookup():

   print("==============================================================================================================")
   print("====INICIALIZACION DE -VLOOKUP-")
   print("==============================================================================================================\n")

   # ----Toma la fecha actual de hoy
   fecha= "{:%Y_%m_%d}".format(datetime.now())

   # ----Se definen los paths de los archivos, el archivo .xlsx y el archivo .csv
   YRA2_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\YRA2_TMOBILE_" + fecha + ".xlsx"
   GIC_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\\F&C GIC - SIG PC List - " + fecha + ".xlsx"

   # ----Se define el nombre y path del documento final
   doc_final_REPORTE_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2','Reporte final') + "\\ Reporte_YRA2_TMOBILE_" + fecha + ".xlsx"
   filename_doc_final_REPORTE_path = "Reporte_YRA2_TMOBILE_" + fecha + ".xlsx"

   # ----Check if file already exists
   directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop','YRA2','Reporte final')
   try:
      os.stat(directorio)
   except:
      os.mkdir(directorio)

   print("========================================================================")
   print("Archivo xlsx = " + YRA2_file_path)
   print("Archivo csv = " + GIC_file_path)
   print("Archivo final = " + doc_final_REPORTE_path)
   print("========================================================================\n")

   print("========================================================================")
   print("----Lectura de los archivos de Reporte YRA2 y GIC como DataFrame de dtype = String")
   print("========================================================================\n")

   # ----Define first DataFrame, especificando que el encabezado empieza en le fila 2 ya que en el codigo de columna la fila 1 es 0 entonces la 3 es 2
   excel_YRA2 = pd.read_excel(YRA2_file_path, header=2, dtype=str) # ----lo que hace es tomar todos los datos en forma de String desde el principio y no arrojaria error ya que va a coincidir con los datos del otro DataFrame
   # ----Define second DataFrame
   excel_GIC = pd.read_excel(GIC_file_path, dtype=str)

   # ----Se renombra el encabezado de la columna AK1=GIC code a AK1=GIC
   excel_YRA2 = excel_YRA2.rename(columns={'GIC code': 'GIC'})
   
   # Escribir en la celda de la fila 0 y la columna 'A'
   #excel_YRA2.at[3, 'GIC'] = 'Eliminar esta Fila'

   # ----No toma las columnas que estan en blanco
   excel_YRA2.columns = excel_YRA2.columns.str.strip()
   excel_YRA2 = excel_YRA2.dropna(how='all')
   excel_YRA2 = excel_YRA2.dropna(axis=1, how='all')
   # ----No toma las columnas que estan en blanco
   excel_GIC.columns = excel_GIC.columns.str.strip()
   excel_GIC = excel_GIC.dropna(how='all')
   excel_GIC = excel_GIC.dropna(axis=1, how='all')    

   print("-------------------------------------------------------------------------")
   print('Columnas del Reporte YRA como DataFrame:\n')
   print(excel_YRA2.columns)
   print("-------------------------------------------------------------------------\n")
   
   # ----Esto convierte los datos de las columnas [['...']] en int (Enteros)
   #excel_YRA2['GIC']=excel_YRA2['GIC'].astype(int)
   # print("--------------------------------")
   # print("Se ejecuto en type str ____ corrigiendo el '.0' a vacio ''")
   # print("--------------------------------\n")
   # ----Esto convierte los datos de las columnas [['...']] en string y en dado caso que tengan '.0' se cambiara por vacio ''
   #excel_YRA2['GIC']=excel_YRA2['GIC'].str.replace(r'\.0+$', '')
   #excel_GIC[['GIC', 'PC Business Group']]=excel_GIC[['GIC', 'PC Business Group']].astype(str)

   print("========================================================================")
   print("----Inicio del vlookup entre reporte YRA2 y archivo GIC")
   print("========================================================================\n")

   print("-------------------------------------------------------------------------")
   print("Datos GIC del Reporte YRA2:\n") 
   print(excel_YRA2['GIC'])
   print("-------------------------------------------------------------------------")
   print("Datos GIC y PC Business Group del Archivo GIC:\n")
   print(excel_GIC[['GIC', 'PC Business Group']])
   print("-------------------------------------------------------------------------\n")

   vlookup_df = pd.merge(excel_YRA2,  
                           excel_GIC[['GIC', 'PC Business Group', 'PC Business Unit', 'PC Business Line']], 
                           on ='GIC', 
                           how ='left')

   # ----View vlookup 
   print("-------------------------------------------------------------------------")
   print("Datos del Vlookup entre el Reporte YRA2 y el archivo GIC:\n" + vlookup_df)
   print("-------------------------------------------------------------------------\n")

   # ----Check if file already exists
   if os.path.isdir(doc_final_REPORTE_path):
         print("========================================================================")
         print('----'+filename_doc_final_REPORTE_path, '____ Exists in the destination path!')
         print("========================================================================\n")
         shutil.rmtree(doc_final_REPORTE_path)

   elif os.path.isfile(doc_final_REPORTE_path):
         os.remove(doc_final_REPORTE_path)
         print("========================================================================")
         print('----'+filename_doc_final_REPORTE_path, '____ Deleted in', 'YRA2', 'becuase is duplicate')
         print("========================================================================\n")

   print("========================================================================")
   print("----Fin del vlookup entre reporte YRA2 y archivo GIC")
   print("========================================================================\n")

   print("========================================================================")
   print("----Inicio de guardado del DataFrame en archivo .XLSX")
   print("----Archivo xlsx final =" + doc_final_REPORTE_path)
   print("========================================================================\n")

   # ----Save vlookup_df to Excel file
   vlookup_df.to_excel(doc_final_REPORTE_path, index=False)

   print("========================================================================")
   print("----Fin de guardado del DataFrame en archivo .XLSX")
   print("----Archivo xlsx final = " + doc_final_REPORTE_path)
   print("========================================================================\n")

   print("==============================================================================================================")
   print("====FINALIZACION DE -VLOOKUP-")
   print("==============================================================================================================\n")


#---------------------------------------- Parte 7 _ Fin del Programa ----------------------------------------
def terminar_programa(connection, hwnd):
        
        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -TERMINAR PROGRAMA-")
        print("==============================================================================================================\n")
        
        print("========================================================================")
        print("----Se termino la automatizacion del Reporte YRA2 -FIN DEL PROGRAMA-")
        print("========================================================================\n")
        
        win= Tk()

        win.attributes('-topmost', True)
        # ----Set the geometry of frame
        win.geometry("400x70")

        #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
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

        #session = None
        #connection = None
        #application = None
        #SapGuiAuto = None
    
        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -TERMINAR PROGRAMA-")
        print("==============================================================================================================\n")

        print("========================================================================")
        print("----Se cerro la conexion de SAP")
        # ----Cierra la pesteña de SAP ejecutada, y solo queda la de Log On
        connection.CloseConnection()
        print("----Se cerro la Pestaña de SAP Logon 770")
        # Envía un mensaje WM_CLOSE a la ventana para cerrarla
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        print("========================================================================\n")

        # ----Sale de ejecutar el PROGRAMA
        exit()


# ----Da los parametros iniciales de la ejecucion de la libreria para ejecutar la pantalla emergente y en su defecto el resto del programa
if __name__ == '__main__':
    
    root = tk.Tk()
    
    # ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    #root.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
    # ----Si se quiere ejecutar en el computador
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
    
    input_form = InputForm(root)
    
    root.mainloop()
