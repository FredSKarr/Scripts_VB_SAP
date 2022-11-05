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

Dim objExcel,objWorkbook
Dim objSheet, Row, i
Set objExcel  = CreateObject("Excel.Application") 
Set objWorkbook = objExcel.Workbooks.open ("C:\SAPScripts\pruebascrip.xlsx")
Set objSheet = objExcel.Worksheets(1)
Row = 1


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nlx02"
session.findById("wnd[0]").sendVKey 0

Do Until objSheet.cells(Row,1).value=""

session.findById("wnd[0]/usr/ctxtS1_LGNUM").text = objSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").text = objSheet.Cells(Row,2).Value
session.findById("wnd[0]/usr/ctxtP_VARI").text = "STDLIST"
session.findById("wnd[0]/usr/ctxtP_VARI").setFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Row = Row + 1
Loop
objExcel.Quit
Set objExcel = Nothing
Set objWorkbook = Nothing
Set objSheet = Nothing
msgbox ("Process Completed")

'------------------------------------------TEST 02--------------------------------------

Dim objExcel
Dim objSheet, intRow, i
Set objExcel = GetObject(,"Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.Activesheet

For i = 2 to objSheet.UsedRange.Rows.Count 'Ciclo de repetición **OJO con UsedRange
'--------------------Excel to sap ------------------------
cOL1 = Trim(CStr(objSheet.Cells(i, 1).Value))'Asignacion de valores a variables para datos
cOL2 = Trim(CStr(objSheet.Cells(i, 2).Value))

'-------------------SAP to Excel ------------------------

objExcel.cells(i, 6).value = session.findById("wnd[0]/sbar").text

'-------------------------Enviar datos a un archivo ------------------------------

aux =cOL1 & " " & cOL2
CreateObject("Wscrip.Shell").run("cmd /c @echo %date %timr " & aux & " >> C:SCRIPT\Log.txt")
next
msgbox"Perfecto"