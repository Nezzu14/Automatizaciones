# Importing the Libraries
import win32com.client
import sys
import subprocess
import time
from datetime import datetime, date
import os
from tkinter import *
from tkinter import messagebox as MessageBox
import convert
import win32gui

def descargar(session, wbs):
        print("DESCARGAR ===========")
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"/DATA"
        try:
           os.stat(directorio)
        except:
           os.mkdir(directorio)
    
        fecha= "{:%Y_%m_%d}".format(datetime.now())

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("0000000003")
        session.findById("wnd[0]/usr/lbl[5,10]").setFocus
        session.findById("wnd[0]/usr/lbl[5,10]").caretPosition = 1
        session.findById("wnd[0]").sendVKey (2)
        session.findById("wnd[0]/usr/lbl[9,13]").setFocus
        session.findById("wnd[0]/usr/lbl[9,13]").caretPosition = 1
        session.findById("wnd[0]").sendVKey (2)
        session.findById("wnd[0]/usr/lbl[13,15]").setFocus
        session.findById("wnd[0]/usr/lbl[13,15]").caretPosition = 0
        session.findById("wnd[0]").sendVKey (2)
        session.findById("wnd[0]/usr/lbl[20,17]").setFocus
        session.findById("wnd[0]/usr/lbl[20,17]").caretPosition = 1
        session.findById("wnd[0]").sendVKey (2)
        session.findById("wnd[0]/usr/chkP_HIER").selected = True
        session.findById("wnd[0]/usr/chkP_ORCVB").selected = True
        session.findById("wnd[0]/usr/ctxtS_PSPID-LOW").text = wbs[0:8]
        session.findById("wnd[0]/usr/ctxtS_POSID-LOW").text = wbs
        session.findById("wnd[0]/usr/chkP_ORCVB").setFocus
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").selectItem ("          1","&Hierarchy")
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem ("          1","&Hierarchy")
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").doubleClickItem ( "          1","&Hierarchy")
        session.findById("wnd[0]/usr/cntlCUSTOM_CONTROL/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCUSTOM_CONTROL/shellcont/shell").selectContextMenuItem ( "&PC")
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directorio
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = wbs+fecha+".htm"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
       # session.findById("wnd[1]").sendVKey (4)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/mbar/menu[1]/menu[0]").select()
        session.findById("wnd[0]/mbar/menu[2]/menu[8]").select()

#puede llegar a fallar el codigo cuando por ejemplo, si el archivo ya existe y se trata de sobrescribir 
# This function will Login to SAP from the SAP Logon window

def saplogin(wbs_list, username,password):
    contraseña= False
    try:

        path = r"C:\Program Files (x86)\SAP\SAPGUI770\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        hwnd = 0
        start_time = time.time()
        while not hwnd:
             hwnd = win32gui.FindWindow(None, 'SAP Logon 770')
             if time.time() - start_time > 30:
                return  # Si se supera el tiempo máximo de espera, se sale de la función
             time.sleep(0.5) 

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        
       # connection = application.OpenConnection("- P20 Production ERP Logistics and Finance", True)
        
        if application.Connections.Count==0 : 
            connection = application.OpenConnection("- P20 Production ERP Logistics and Finance", True)
            session = connection.Sessions(0)
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

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
      

       # session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

  

     
        
        wbs_list=  [elemento for elemento in wbs_list if elemento]#eliminar posibles espacios vacios
        
        if len(wbs_list)>0:

           for wbs in wbs_list:
               descargar(session, wbs)
               print(wbs)
               print("WB CARGANDOOOOOOO")
         
     
        if len(wbs_list)>0:
              convert.transformar_fichero() 
              convert.unir()
    

   

    except:
        print(sys.exc_info())
        
        win= Tk()

        win.attributes('-topmost', True)
        #Set the geometry of frame
        win.geometry("600x250")
        
        def close_win():
           win.destroy()
        
        #Create a text label
        Label(win,text=sys.exc_info(), font=('Helvetica',10)).pack(pady=20)
        
        #Create Entry Widget for password
        
        #Create a button to close the window
        Button(win, text="Quit", font=('Helvetica bold',
        10),command=close_win).pack(pady=20, side="top")
        
        win.mainloop()

    


    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None
     
    return contraseña
    
    exit()
       