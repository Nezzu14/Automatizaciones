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
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "migumart"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "********"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
