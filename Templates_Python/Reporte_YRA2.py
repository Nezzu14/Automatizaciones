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



# ======================================== Parte 1 _ Inicializacion ========================================
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
# ==========================================================================================================



# ======================================== Parte 2 _ Ejecucion SAP y el resto de Procesos ========================================
#---------------------------------------------------------------------------------------------------------------------------------
def callback(hwnd, hwnds):
    # ----Definir la función de callback para EnumWindows de SAP LOGON 770
    if 'SAP Logon' in win32gui.GetWindowText(hwnd):
        hwnds.append(hwnd)

def saplogin(username, password, radate, createdon, to, wbs_list):

    # ----Crear una lista para almacenar los identificadores de ventana
    hwnds = []
    # ----Llamar a EnumWindows para enumerar todas las ventanas y agregar los identificadores de ventana a la lista
    win32gui.EnumWindows(callback, hwnds)
    # ----Obtener la cantidad de ventanas hijas
    n_windows = len(hwnds)
    # ----Imprimir la cantidad de ventanas hijas
    
    print("------------------------")
    print("SAP LOGON Antes de if == 0: ", n_windows)
    print("------------------------\n")

    if n_windows == 0:
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
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
        except:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -ERROR DE SAP-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----Error de SAP: ")
            # ----Este print es para que me muestre el Error en especifico si no lo se: print(sys.exc_info()), ya que de esta forma se sabe que el error generado es "pywintypes.com_error"
            print(sys.exc_info())
            print("========================================================================\n")

            win= Tk()

            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("510x230")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - DATOS INCORRECTOS")

            def close_win():
                win.destroy()

            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. Error de SAP', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Una vez cerrada la ventana de error de SAP y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)

            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=10, side="top")

            win.mainloop()

            #print("========================================================================")
            #print("----Se cerro la Pestaña de SAP Logon 770")
            # ----Envía un mensaje WM_CLOSE a la ventana para cerrarla
            #win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            #print("========================================================================\n")

            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
            print("==============================================================================================================\n")

            # ----Sale de ejecutar el PROGRAMA PYTHON
            exit()

    # ----Crear una lista para almacenar los identificadores de ventana
    hwnds = []
    # ----Llamar a EnumWindows para enumerar todas las ventanas y agregar los identificadores de ventana a la lista
    win32gui.EnumWindows(callback, hwnds)
    # ----Obtener la cantidad de ventanas hijas
    n_windows = len(hwnds)
    # ----Imprimir la cantidad de ventanas hijas

    print("------------------------")
    print("SAP LOGON Antes de if == 1: ", n_windows)
    print("------------------------\n")

    if n_windows == 1:
        hwnd = win32gui.FindWindow(None, 'SAP Logon 770')

        # ----Detecta la ventana de SAP
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        try:
            # ----Se ejecuta la Connection del ERP de SAP P20 de cero
            connection = application.OpenConnection("- P20 Production ERP Logistics and Finance", True)
        except pywintypes.com_error as e:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -ERROR DE SAP-")
            print("==============================================================================================================\n")
    
            print("========================================================================")
            print("----Error de SAP: ")
            # ----Este print es para que me muestre el Error en especifico si no lo se: print(sys.exc_info()), ya que de esta forma se sabe que el error generado es "pywintypes.com_error"
            print(sys.exc_info())
            print("========================================================================\n")
    
            win= Tk()
    
            win.attributes('-topmost', True)
            # ----Set the geometry of frame
            win.geometry("510x230")
            #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
            # ----Si se quiere ejecutar en el computador
            win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
            win.title("REPORTE YRA2 - DATOS INCORRECTOS")
    
            def close_win():
                win.destroy()
    
            # ----Create a text label
            Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
            Label(win,text='1. No se encuentra conectado a la VPN de CISCO o la red de Nokia', font=('Helvetica',10,'bold')).pack(pady=1)
            Label(win,text='= Conectese a la VPN de CISCO o la red de Nokia y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
            Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
            Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
    
            # ----Create a button to close the window
            Button(win, text="Quit", font=('Helvetica bold',
            10),command=close_win).pack(pady=10, side="top")
    
            win.mainloop()
    
            print("========================================================================")
            print("----Se cerro la Pestaña de SAP Logon 770")
            # ----Envía un mensaje WM_CLOSE a la ventana para cerrarla
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            print("========================================================================\n")
    
            print("==============================================================================================================")
            print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
            print("==============================================================================================================\n")
    
            # ----Sale de ejecutar el PROGRAMA PYTHON
            exit()

        try:
            # ----Se define la "session" de la "connection" para poder interactuar con la pagina de SAP
            session = connection.Sessions(0)

            # ----Ingreso de Usuario y Contraseña
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            session.findById("wnd[0]").sendVKey(0)
 
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # ----Obtener información de la sesión
            session_info = session.Info
            usuario=str(session_info.User)
            print("------------------------")
            print("----Lenght Usuario:",str(len(usuario)))
            print("------------------------\n")

            # ----Verificar si el inicio de sesión fue exitoso
            if len(usuario) == 0:
                print("==============================================================================================================")
                print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -REPORTE YRA2_DATOS INCORRECTOS-")
                print("==============================================================================================================\n")
                print("========================================================================")
                print("----Se ingresaron los datos de usuario y/o contraseña de forma incorrecta")
                print(sys.exc_info())
                print("========================================================================\n")

                win= Tk()

                win.attributes('-topmost', True)
                # ----Set the geometry of frame
                win.geometry("500x240")
                #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
                # ----Si se quiere ejecutar en el computador
                win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
                win.title("REPORTE YRA2 - DATOS INCORRECTOS")

                def close_win():
                    win.destroy()

                # ----Create a text label
                Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
                Label(win,text='1. Usuario y/o Contraseña incorrectos', font=('Helvetica',10,'bold')).pack(pady=1)
                Label(win,text='= Ejecute el programa de nuevo e ingrese los datos correctamente\n', font=('Helvetica',10)).pack(pady=0.1)
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
                #print("----Se cerro la Pestaña de SAP Logon 770")
                # ----Envía un mensaje WM_CLOSE a la ventana para cerrarla
                #win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
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
             
            print("==============================================================================================================")
            print("====FINALIZACION DE -SAP LOGIN-")
            print("==============================================================================================================\n")  
        except pywintypes.com_error:
            # ----Obtener información de la sesión
            session_info = session.Info
            usuario=str(session_info.User)
            print("------------------------")
            print("----Lenght Usuario:",str(len(usuario)))
            print("------------------------\n")

            # ----Verificar si el inicio de sesión fue exitoso
            if len(usuario) == 0:
                print("==============================================================================================================")
                print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -REPORTE YRA2_DATOS INCORRECTOS-")
                print("==============================================================================================================\n")
                print("========================================================================")
                print("----Se ingresaron los datos de usuario y/o contraseña de forma incorrecta")
                print(sys.exc_info())
                print("========================================================================\n")

                win= Tk()

                win.attributes('-topmost', True)
                # ----Set the geometry of frame
                win.geometry("500x240")
                #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
                # ----Si se quiere ejecutar en el computador
                win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
                win.title("REPORTE YRA2 - DATOS INCORRECTOS")

                def close_win():
                    win.destroy()

                # ----Create a text label
                Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
                Label(win,text='1. Usuario y/o Contraseña incorrectos', font=('Helvetica',10,'bold')).pack(pady=1)
                Label(win,text='= Ejecute el programa de nuevo e ingrese los datos correctamente\n', font=('Helvetica',10)).pack(pady=0.1)
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
                #print("----Se cerro la Pestaña de SAP Logon 770")
                # ----Envía un mensaje WM_CLOSE a la ventana para cerrarla
                #win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
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
             
            print("==============================================================================================================")
            print("====FINALIZACION DE -SAP LOGIN-")
            print("==============================================================================================================\n")
    else:
        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -DOS O MAS LOGON 770 ABIERTOS-")
        print("==============================================================================================================\n")

        print("========================================================================")
        print("----Error de SAP LOGON 770: ")
        # ----Este print es para que me muestre el Error en especifico si no lo se: print(sys.exc_info()), ya que de esta forma se sabe que el error generado es "pywintypes.com_error"
        print(sys.exc_info())
        print("========================================================================\n")

        win= Tk()

        win.attributes('-topmost', True)
        # ----Set the geometry of frame
        win.geometry("520x220")
        #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
        # ----Si se quiere ejecutar en el computador
        win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
        win.title("REPORTE YRA2 - DATOS INCORRECTOS")

        def close_win():
            win.destroy()

        # ----Create a text label
        Label(win,text='\nSE HA PRODUCIDO UN ERROR POR ESTA RAZON:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
        Label(win,text='1. Dos o mas ventanas SAP LOGON 770 abiertas', font=('Helvetica',10,'bold')).pack(pady=1)
        Label(win,text='= Una vez cerradas las ventanas de SAP LOGON 770 y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
        Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
        Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)

        # ----Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=10, side="top")

        win.mainloop()

        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DOS LOGON 770 ABIERTOS-")
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
                if i < 5:
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," + str(i) + "]").text = str(element)
                
                else:
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," + "5" + "]").text = str(element)
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = str(i-4)
                
                print("========================================================================")
                print("Linea en la columna y WBS de linea en la columna: ", i, element)
                print("Posicion del Scrollbar:", str(i))
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
        except pywintypes.com_error as e_sap:
            print("==============================================================================================================")
            print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -DATOS INGRESADOS INCORRECTOS EN YRA2-")
            print("==============================================================================================================\n")

            print("========================================================================")
            print("----Se ingresaron los datos de forma incorrecta de la transaccion YRA2")
            print(str(e_sap))
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
            #print("----Se cerro la Pestaña de SAP Logon 770")
            # ----Envía un mensaje WM_CLOSE a la ventana para cerrarla
            #win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
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
#---------------------------------------------------------------------------------------------------------------------------------



# ======================================== CONFIGURACION PARA LA EJECUCION DEL PROMARAMA ========================================
# ----Da los parametros iniciales de la ejecucion de la libreria para ejecutar la pantalla emergente y en su defecto el resto del programa
if __name__ == '__main__':
    
    root = tk.Tk()
    
    # ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    #root.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
    # ----Si se quiere ejecutar en el computador
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
    
    input_form = InputForm(root)
    
    root.mainloop()
# ===============================================================================================================================