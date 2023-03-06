if application.Connections.Count==0 : 
    connection = application.OpenConnection("- P20 Production ERP Logistics and Finance", True)
    session = connection.Sessions(0)
    # ----Ingreso de Usuario y Contraseña
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]").sendVKey(0
else: 
    if application.Connections.Count<6:
          connection= application.Connections(0)
          session = connection.Sessions(0)
          session.CreateSession()
          session=connection.Sessions(connection.Sessions.Count -1)
    else:
        print("Couldn't connect to application because sap reach the maximum number of sessions"
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