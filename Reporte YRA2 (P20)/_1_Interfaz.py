import tkinter as tk
import _2_Open_sap
import pickle

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

        # ----Load saved username and password if they exist
        try:
            with open('login_info.bin', 'rb') as f:
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
        print(variante)

        # ----Close the window and end the program pero si quieren seguir las varialbles se debe pner return al final del todo
        self.master.destroy()
        
        print("========================================================================================================================")
        print("====FINALIZACION DE LA VENTANA EMERGENTE DE -DESCARGA YRA2-")
        print("========================================================================================================================\n")

        # ----Save the login information
        self.login_info['username'] = username
        self.login_info['password'] = password
       
        with open('login_info.bin', 'wb') as f:
            pickle.dump(self.login_info, f)

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA LA APERTURA DE SAP
        _2_Open_sap.saplogin(variante, username, password) 
        #--------------------------------------------------------------------------------------------------------------------

        print("==============================================================================================================")
        print("====FINALIZACION DE -REPORTE YRA2-")
        print("==============================================================================================================\n")

        # ----Esto permite que se cierre la ventana emergente con self.master.destroy() pero que las variables no se borren
        return

# ----Da los parametros iniciales de la ejecucion de la libreria para ejecutar la pantalla emergente
if __name__ == '__main__':
    
    root = tk.Tk()
    
    # ----La "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\Archivos personales\\Automatizacion Python\\Reporte YRA2 (P20)\\nokia.ico")
    
    input_form = InputForm(root)
    
    root.mainloop()
