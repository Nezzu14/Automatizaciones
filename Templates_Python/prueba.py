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

    def __init__(self, master):

        print("==============================================================================================================")
        print("====INICIALIZACION DE -REPORTE YRA2-")
        print("==============================================================================================================\n")

        print("==============================================================================================================")
        print("====INICIALIZACION DE LA VENTANA EMERGENTE DE -DESCARGA YRA2-")
        print("==============================================================================================================\n")

        # ----Genera el titulo de la pantalla emergente
        self.master = master
        master.title('INICIO DESCARGA YRA2') #Titulo en el Pop up de ingresar Usuario y contraeña

        # ----El directorio de los datos incriptados del usuario y password
        directorio_login_bin = (r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\login_info.bin")
            # ----Si se quiere ejecutar en el computador  
        #directorio_login_bin = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Nokia', 'Archivos personales', 'Automatizacion Python', 'Reporte YRA2 (P20)') + '\\login_info.bin'

        # ----Load saved username and password if they exist
        try:
            with open(directorio_login_bin, 'rb') as f:
                self.login_info = pickle.load(f)
        except:
            self.login_info = {'username': '', 'password': ''}

        # ----Create labels and input fields
        self.username_label = tk.Label(master, text='Username:')
        self.username_label.grid(row=0, column=0, padx=5, pady=5)
        self.username_input = tk.Entry(master)
        self.username_input.insert(0, self.login_info['username'])
        self.username_input.grid(row=0, column=1, padx=5, pady=5)

        self.password_label = tk.Label(master, text='Password:')
        self.password_label.grid(row=1, column=0, padx=5, pady=5)
        self.password_input = tk.Entry(master, show='*')
        self.password_input.insert(0, self.login_info['password'])
        self.password_input.grid(row=1, column=1, padx=5, pady=5)

        self.variante_label = tk.Label(master, text='Variante (Plantilla - Modified by):')
        self.variante_label.grid(row=2, column=0, padx=5, pady=5)
        self.variante_input = tk.Entry(master)
        self.variante_input.grid(row=2, column=1, padx=5, pady=5)

        self.wbs_label = tk.Label(master, text='WBS (Separar con comas):')
        self.wbs_label.grid(row=3, column=0, padx=5, pady=5)
        self.wbs_input = tk.Text(master, height=5, width=15)
        self.wbs_input.grid(row=3, column=1, padx=5, pady=5)

        # ----Create submit button
        self.submit_button = tk.Button(master, text='Submit', command=self.submit)
        self.submit_button.grid(row=4, column=1, padx=5, pady=5)


    def submit(self):

        print("========================================================================")
        print("----Se presionó el boton SUBMIT ____ Procede a inicar -open_sap.saplogin(variante,username, password)-")
        print("========================================================================\n")

        # ----Get the values of the input fields and do something with them
        username = self.username_input.get()
        password = self.password_input.get()
        variante = self.variante_input.get()

        wbs = self.wbs_input.get('1.0', 'end-1c')
        print(wbs)
        #wbs_list = wbs.split("\n")or(",")or(" ")or(", ")        
        
        print("========================================================================")
        print("Username = " + username)
        print("Password = " + password)
        print("Variante = " + variante)
        #print("WBS List = ", wbs_list)
        print("========================================================================\n")
        # ----Close the window and end the program pero si quieren seguir las varialbles se debe pner return al final del todo
        self.master.destroy()
        
        print("========================================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DESCARGA YRA2-")
        print("========================================================================================================================\n")

        pyperclip.copy(wbs)
        print(pyperclip)

        # ----Save the login information
        self.login_info['username'] = username
        self.login_info['password'] = password
       
        # ----El directorio de los datos incriptados del usuario y password
        directorio_login_bin = (r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\login_info.bin")
            # ----Si se quiere ejecutar en el computador        
        #directorio_login_bin = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Nokia', 'Archivos personales', 'Automatizacion Python', 'Reporte YRA2 (P20)') + '\\login_info.bin'
        
        with open(directorio_login_bin, 'wb') as f:
            pickle.dump(self.login_info, f)



# ----Da los parametros iniciales de la ejecucion de la libreria para ejecutar la pantalla emergente y en su defecto el resto del programa
if __name__ == '__main__':
    
    root = tk.Tk()
    
    # ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    #root.iconbitmap(r"C:\Program Files (x86)\Nokia\Reporte YRA2\Reporte_YRA2\nokia.ico")
    # ----Si se quiere ejecutar en el computador
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
    
    input_form = InputForm(root)
    
    root.mainloop()