If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "12345"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btnB_EXPACC").press
session.findById("wnd[0]/usr/btn%_S_RACAT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "123"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/chkP_CUSPO").selected = true
session.findById("wnd[0]/usr/chkP_CUSMNO").selected = true
session.findById("wnd[0]/usr/radP_ITEM").select
session.findById("wnd[0]/usr/chkP_NO0ALL").selected = true
session.findById("wnd[0]/usr/ctxtP_VARI").text = "123456"
session.findById("wnd[0]/usr/ctxtP_VARI").setFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
