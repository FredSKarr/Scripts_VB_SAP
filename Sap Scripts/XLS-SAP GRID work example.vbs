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

'Select Menu
session.findById("wnd[0]/mbar/menu[5]/menu[2]/menu[1]").select
'Set Varibale with the GRID
set tbl = session.findById("wnd[1[/usr/sscubD0500_SUBSCREEN_SAPLSLCV_DIALOG:0501/cntG51_CONTAINER/shellcont/shell")
'Ciclo para encontar una coincidencia
tbl.setFocus
WScript.Sleep(1000)

For i = 0 To tbl.RowCount - 1
    If tbl.GetCellValue(CInt(i),"VARIANT") = "/ALE1" Then 'buscar la conincidencia
        tbl.currentCellRow = CInt(i) 'si la encuentra asignarla a la variable i
        tbl.selectedRows = CInt(i) 'Seleccionar la fila
        tbl.clickCurrentCell 'dar click a la celda
        
        Exit For
    End If
Next

msgbox "Done", vbInfomation