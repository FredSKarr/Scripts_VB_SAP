Dim App, Connection, session As Object
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

'launch a transaction
session.findById("wnd[0]").Maximize

'Metodo 1
session.findById("wnd[0]/tbar[0]/okcd").Text = "FS10N" 'TCODE
session.findById("wnd[0]").sendVKey 0

'Metodo 2
session.StartTransaction("FS10N")

