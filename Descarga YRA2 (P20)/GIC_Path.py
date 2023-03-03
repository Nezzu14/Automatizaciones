#SEGUNDO EN CORRER DE LOS ARCHIVOS GIC

import tkinter as tk
import pickle
import GIC_Mover


class Archivo_GIC:
    def __init__(self, master):
        
        #   Increase window size
        #self.master_width = 1500
        #self.master_height = 700
        
        self.master = master
        #master.configure(width=self.master_width, height=self.master_height)
        master.title('Path del Archivo de GIC') #Titulo en el Pop up de ingresar Usuario y contraeña

        # Load saved Path if they exist
        try:
            with open('Path_GIC.bin', 'rb') as f:
                self.Path_GIC = pickle.load(f)
        except:
            self.Path_GIC = {'Path': '', 'Nombre_GIC': ''}

        # Create labels and input fields
        self.Path_label = tk.Label(master, text='Click a "Conservar" en el Pop up de la descarga')
        self.Path_label.grid(row=0, column=0, padx=30, pady=10)
        

        self.Path_label = tk.Label(master, text='Path - Direccion del archivo (Downloads o Descargas):      ')
        self.Path_label.grid(row=2, column=0, padx=30, pady=10)
        self.Path_input = tk.Entry(master)
        self.Path_input.insert(0, self.Path_GIC['Path'])
        self.Path_input.grid(row=2, column=1, padx=30, pady=10)

        self.Nombre_GIC_label = tk.Label(master, text='Nombre_GIC:      ')
        self.Nombre_GIC_label.grid(row=4, column=0, padx=30, pady=10)
        self.Nombre_GIC_input = tk.Entry(master)
        self.Nombre_GIC_input.insert(0, self.Path_GIC['Nombre_GIC'])
        self.Nombre_GIC_input.grid(row=4, column=1, padx=30, pady=10)

        # Create submit button
        self.submit_button = tk.Button(master, text='Submit', command=self.submit)
        self.submit_button.grid(row=6, column=1, padx=5, pady=5)

    def submit(self):
        # Get the values of the input fields and do something with them
        Path = self.Path_input.get()
        Nombre_GIC = self.Nombre_GIC_input.get()
        print("========================================================================")
        print(Path) 
        print(Nombre_GIC)


        #SE EJECUTA MOVER EL GIC
        GIC_Mover.Mover_GIC(Nombre_GIC)


        # Save the login information
        self.Path_GIC['Path'] = Path
        self.Path_GIC['Nombre_GIC'] = Nombre_GIC
       
        with open('Path_GIC.bin', 'wb') as f:
            pickle.dump(self.Path_GIC, f)

#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#if __name__ == '__main__':
#    root = tk.Tk()
#    root.iconbitmap(r"C:\Users\migumart\OneDrive - Nokia\Archivos personales\Automatizacion Python\Descarga YRA2 (P20)\nokia.ico")
#    # el "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
#    input_form = Archivo_GIC(root)
#    root.mainloop()

