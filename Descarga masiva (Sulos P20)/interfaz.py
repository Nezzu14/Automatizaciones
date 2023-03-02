import tkinter as tk
import open_sap
import pickle

class InputForm:
    def __init__(self, master):
        self.master = master
        master.title('SULOS 1.0 ALFA') #Titulo en el Pop up de ingresar Usuario y contraeña

        # Load saved username and password if they exist
        try:
            with open('login_info.bin', 'rb') as f:
                self.login_info = pickle.load(f)
        except:
            self.login_info = {'username': '', 'password': ''}

        # Create labels and input fields
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

        self.text_label = tk.Label(master, text='WBS (Separar con comas):')
        self.text_label.grid(row=2, column=0, padx=5, pady=5)
        self.text_input = tk.Text(master, height=5, width=30)
        self.text_input.grid(row=2, column=1, padx=5, pady=5)

        # Create submit button
        self.submit_button = tk.Button(master, text='Submit', command=self.submit)
        self.submit_button.grid(row=3, column=1, padx=5, pady=5)

    def submit(self):
        # Get the values of the input fields and do something with them
        username = self.username_input.get()
        password = self.password_input.get()
        wbs = self.text_input.get('1.0', 'end-1c')
        wbs_list = wbs.split("\n")or(",")or(" ")or(", ")
        print(wbs_list)
        open_sap.saplogin(wbs_list,username, password)

        # Save the login information
        self.login_info['username'] = username
        self.login_info['password'] = password
        with open('login_info.bin', 'wb') as f:
            pickle.dump(self.login_info, f)

if __name__ == '__main__':
    root = tk.Tk()
    root.iconbitmap(r"C:\Users\migumart\OneDrive - Nokia\Archivos personales\Automatizacion Python (Sulos)\Descarga masiva\nokia.ico")
    # el "r" es para que el path de la imagen no tome como caracteres especiales los slash "\" sino como texto
    input_form = InputForm(root)
    root.mainloop()
