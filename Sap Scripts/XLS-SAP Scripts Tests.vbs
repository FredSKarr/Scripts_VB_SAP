Sub prueba3()
Dim App, Connection, session As Object
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)


On Error GoTo Errhandler

    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    Range("I" & IntContador).FormulaR1C1 = session.findById("wnd[0]/sbar").Text

Errhandler:

      ' If an error occurs, display a message and end the macro.
      MsgBox "An error has occurred. The macro will end."
