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
        self.radate_input = tk.Text(master, height=1, width=12,)
        self.radate_input.grid(row=3, column=1, padx=5, pady=5)
        self.radate_input.bind('<KeyRelease>', lambda event: self.on_entry_changed(self.radate_input.get()[:10]))

        self.createdon_label = tk.Label(master, text='Created on:')
        self.createdon_label.grid(row=4, column=0, padx=5, pady=5)
        self.createdon_input = tk.Text(master, height=1, width=12)
        self.createdon_input.grid(row=4, column=1, padx=5, pady=5)
        self.createdon_input.bind('<KeyRelease>', lambda event: self.on_entry_changed(self.createdon_input.get()))

        self.to_label = tk.Label(master, text='To:')
        self.to_label.grid(row=4, column=2, padx=5, pady=5)
        self.to_input = tk.Text(master, height=1, width=12)
        self.to_input.grid(row=4, column=3, padx=5, pady=5)
        self.to_input.bind('<KeyRelease>', lambda event: self.on_entry_changed(self.to_input.get()))

        self.wbs_label = tk.Label(master, text='WBS\n(Separar por "Saltos de Linea"):')
        self.wbs_label.grid(row=5, column=0, padx=5, pady=5)
        self.wbs_input = tk.Text(master, height=5, width=12)
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
        #wbs_list = wbs.split("\n")or(",")or(" ")or(", ")        

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
        saplogin(username, password, radate, createdon, to, wbs) 
        #--------------------------------------------------------------------------------------------------------------------

        print("========================================================================")
        print("----Una vez guardado el reporte YRA2 en .xls, se corregira en .xlsx con -Correccion_formato._3_Deshabiiltar_error()-")
        print("========================================================================\n")

        # ----Esto permite que se cierre la ventana emergente con self.master.destroy() pero que las variables no se borren
        return


#---------------------------------------- Parte 2 _ Open SAP ----------------------------------------
def saplogin(username, password, radate, createdon, to, wbs):

    # ----This function will Login to SAP from the SAP Logon window
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
        else: 
            # ----Si ya hay sesiones abiertas con el acceso del usuario entonces se abrira una sesion aparte y se empezara a ejecitar el proceso de descarga de YRA2
            # ----Se abrira solo hasta el maximo de 6 sesiones, si ya hay 6 sesiones abiertas entonces arrojara un error, el cual es el de "except"
            if application.Connections.Count<6:
                  connection= application.Connections(0)
                  session = connection.Sessions(0)
                  session.CreateSession()
                  session=connection.Sessions(connection.Sessions.Count -1)
            else:
                print("Couldn't connect to application because sap reach the maximum number of sessions")

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

        username= username
        
        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA DESCARGA DEL REPORTE YRA2 Y DEL ARCHIVO GIC
        Path_YRA2_SAP(session, username, radate, createdon, to, wbs)
        #--------------------------------------------------------------------------------------------------------------------
        
    except:

        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -REPORTE YRA2_DATOS INCORRECTOS-")
        print("==============================================================================================================\n")

        print("========================================================================")
        print("----Se ingresaron los datos de usuario y/o contraseña de forma incorrecta")
        print("========================================================================\n")

        print(sys.exc_info())
        
        win= Tk()

        win.attributes('-topmost', True)
        # ----Set the geometry of frame
        win.geometry("500x350")
        #win.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
        # ----Si se quiere ejecutar en el computador
        win.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
        win.title("REPORTE YRA2 - DATOS INCORRECTOS")

        def close_win():
           win.destroy()
        
        # ----Create a text label
        Label(win,text='\nSE HA PRODUCIDO UN ERROR POR UNA DE ESTAS DOS RAZONES:\n', font=('Helvetica',10,'italic')).pack(pady=0.1)
        Label(win,text='1. Usuario, Contraseña y/o demas datos incorrectos', font=('Helvetica',10,'bold')).pack(pady=1)
        Label(win,text='= Ejecute el programa de nuevo e ingrese los datos correctamente\n', font=('Helvetica',10)).pack(pady=0.1)
        Label(win,text='2. Tiene seis sesiones abiertas, el cual es el maximo para SAP', font=('Helvetica',10,'bold')).pack(pady=1)
        Label(win,text='= Cierre una de esas seis sesiones y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
        Label(win,text='3. No se encuentra conectado a la VPN de CISCO o la red de Nokia', font=('Helvetica',10,'bold')).pack(pady=1)
        Label(win,text='= Conectese a la VPN de CISCO o la red de Nokia y vuelva a ejecutar el programa\n', font=('Helvetica',10)).pack(pady=0.1)
        Label(win,text='--> Para volver a ejecutar el programa <--', font=('Helvetica',10,'bold','underline')).pack(pady=1)
        Label(win,text='* Darle a "Quit" y vuelva a iniciar el programa *', font=('Helvetica',10)).pack(pady=0.1)
 
        # ----Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=10, side="top")
        
        win.mainloop()

        print("==============================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DATOS INCORRECTOS-")
        print("==============================================================================================================\n")

                # ----Sale de ejecutar el PROGRAMA
        exit()
    
    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None

    #==============================================================================================================
    #====FINALIZACION DE -SAP LOGIN- \\\\CODIGO
    #==============================================================================================================    

def Path_YRA2_SAP(session, username, radate, createdon, to, wbs):
        
        print("==============================================================================================================")
        print("====INICIALIZACION DE -PATH YRA2 SAP-")
        print("==============================================================================================================\n")  

        print("========================================================================")
        print(username)
        print(radate)
        print(createdon)
        print(to)
        print(wbs)
        print("========================================================================\n")

        username= username

        # ----Indicativo de la fecha actual
        fecha= "{:%Y_%m_%d}".format(datetime.now())

        print("========================================================================")
        print("----Entrando a YRA2 en SAP")
        print("========================================================================\n")

        # ----Check if file already exists
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"/YRA2"
        try:
           os.stat(directorio)
        except:
           os.mkdir(directorio)
        
        # ----Aca inicia el script the SAP hecho por SAP y se ejecuta entrando  a la transaccion de YRA2
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "yra2"
        session.findById("wnd[0]").sendVKey (0)
        #session.findById("wnd[0]/tbar[1]/btn[17]").press()

        session.findById("wnd[0]/usr/ctxtS_PSPID-LOW").text = "G-0609A3EZ1" # ----Top WBS por defecto para que deje abrir la ventana emergente y pegar todas las WBS
        session.findById("wnd[0]/usr/ctxtP_DATUM").text = radate # ----RA Date
        session.findById("wnd[0]/usr/ctxtP_DATUM").caretPosition = 10
        session.findById("wnd[0]/usr/btn%_S_PSPID_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()


# ----Da los parametros iniciales de la ejecucion de la libreria para ejecutar la pantalla emergente y en su defecto el resto del programa
if __name__ == '__main__':
    
    root = tk.Tk()
    
    # ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    #root.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
    # ----Si se quiere ejecutar en el computador
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
    
    input_form = InputForm(root)
    
    root.mainloop()
