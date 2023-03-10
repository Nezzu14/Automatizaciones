import tkinter as tk
import pickle
import os
import _2_Open_sap
import _3_Correccion_formato_xlsx
import _4_GIC_Descarga
import _5_Correccion_formato_csv
import _6_Clasificador
import _7_Fin_del_programa


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
        directorio_login_bin = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Nokia', 'Archivos personales', 'Automatizacion Python', 'Reporte YRA2 (P20)') + '\\login_info.bin'

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

        # ----Create submit button
        self.submit_button = tk.Button(master, text='Submit', command=self.submit)
        self.submit_button.grid(row=3, column=1, padx=5, pady=5)


    def submit(self):

        print("========================================================================")
        print("----Se presionó el boton SUBMIT ____ Procede a inicar -open_sap.saplogin(variante,username, password)-")
        print("========================================================================\n")

        # ----Get the values of the input fields and do something with them
        username = self.username_input.get()
        password = self.password_input.get()
        variante = self.variante_input.get()
        print("========================================================================")
        print("Variante = " + variante)
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
        directorio_login_bin = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Nokia', 'Archivos personales', 'Automatizacion Python', 'Reporte YRA2 (P20)') + '\\login_info.bin'
        with open(directorio_login_bin, 'wb') as f:
            pickle.dump(self.login_info, f)

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA LA APERTURA DE SAP
        _2_Open_sap.saplogin(variante, username, password) 
        #--------------------------------------------------------------------------------------------------------------------

        print("========================================================================")
        print("----Una vez guardado el reporte YRA2 en .xls, se corregira en .xlsx con -Correccion_formato._3_Deshabiiltar_error()-")
        print("========================================================================\n")

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE CORREGIRA EL FORMATO DEL REPORTE YRA2
        _3_Correccion_formato_xlsx.Deshabiiltar_error()
        #--------------------------------------------------------------------------------------------------------------------

        print("========================================================================")
        print("----Una vez guardado correctamente el reporte YRA2, se empieza a ejecutar como segundo proceso -_4_GIC_Descarga.Descargar GIC-")
        print("========================================================================\n")
    
        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EMPEZARA A DESCARGAR EL ARCHIVO GIC
        _4_GIC_Descarga.Descargar_GIC()
        #--------------------------------------------------------------------------------------------------------------------

        print("========================================================================")
        print("----Una vez descargado el archivo GIC, se empieza a corregir el formato del archivo .csv -_5_GIC_Descarga.Descargar GIC-")
        print("========================================================================\n")
    
        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EMPEZARA A CORREGIR EL ARCHIVO GIC
        _5_Correccion_formato_csv.cambio_formato_csv()
        #--------------------------------------------------------------------------------------------------------------------

        print("========================================================================")
        print("----Una vez guardado correctamente el reporte YRA2 y el archivo GIC, se empieza a ejecutar el Vlookup entre ambos archivos -_6_Clasificador.vlookup-")
        print("========================================================================\n")

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTARA EL VLOOKUP ENTRE EL ARCHIVO .CSV A .XLSX
        _6_Clasificador.vlookup()
        #-------------------------------------------------------------------------------------------------------------------- 

        print("==============================================================================================================")
        print("====FINALIZACION DE -REPORTE YRA2-")
        print("==============================================================================================================\n")

        print("========================================================================")
        print("----Una vez finalizado el proceso del Reporte YRA2 se terminara el programa -_7_Fin_del_programa.terminar_programa-")
        print("========================================================================\n")

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTARA EL VLOOKUP ENTRE EL ARCHIVO .CSV A .XLSX
        _7_Fin_del_programa.terminar_programa()
        #-------------------------------------------------------------------------------------------------------------------- 

        # ----Esto permite que se cierre la ventana emergente con self.master.destroy() pero que las variables no se borren
        return



# ----Da los parametros iniciales de la ejecucion de la libreria para ejecutar la pantalla emergente
if __name__ == '__main__':
    
    root = tk.Tk()
    
    # ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
    
    input_form = InputForm(root)
    
    root.mainloop()
