Sub PrintTOs()

Dim App, Connection, session As Object
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)
Dim IntUltimaF, IntTONumber, IntItem, IntContador As Integer 'variables ultima fila con datos, Numero de TO y numero de item
Dim StrPrinter As String 'variable impresora
Dim Resp As Byte

'Asignación de valor a variables iniciales
IntContador = 5
IntUltimaF = (Range("B5").End(xlDown).Offset(1, 0).Row)
Select Case Range("X1").Text
Case Is = 1
    StrPrinter = "YGDV"
Case Is = 2
    StrPrinter = "YGG0"
End Select

'launch a transaction
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nLT31" 'TCODE LT31 (PrintTO´s)
session.findById("wnd[0]").sendVKey 0


Do While IntContador <> 0
    'Asignación de variables ciclicas
    IntTONumber = Range("B" & IntContador).Text
    IntItem = Range("C" & IntContador).Text
    ' transaction parameters and commands
    session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").Text = "020" 'WH number
    session.findById("wnd[0]/usr/txtLTAK-TANUM").Text = IntTONumber ' TO number
    session.findById("wnd[0]/usr/txtRL03T-TAPOS").Text = IntItem ' item number
    session.findById("wnd[0]/usr/ctxtRLDRU-DRUKZ").Text = "02" ' print code
    session.findById("wnd[0]/usr/ctxtRLDRU-LDEST").Text = StrPrinter 'Printer
    session.findById("wnd[0]/usr/ctxtRLDRU-LDEST").SetFocus
    session.findById("wnd[0]/usr/ctxtRLDRU-LDEST").caretPosition = 4
    session.findById("wnd[0]/tbar[0]/btn[86]").press 'Boton imprimir
    IntContador = IntContador + 1 'incremental
    If IntContador = IntUltimaF Then Exit Do 'salida del ciclo
Loop
    session.findById("wnd[0]").sendVKey 3 'back

Resp = MsgBox("¿Limpiar datos?", vbQuestion + vbYesNo, "Impresión de TO´s")  'opción para limpieza de datos

If Resp = vbYes Then
IntUltimaF = IntUltimaF + 1
Range("A5:P" & IntUltimaF).ClearContents
Range("A5:P" & IntUltimaF).Select
    With Selection.Interior
        .PatternColor = 255
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDashDotDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDashDotDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDashDotDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDashDotDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

Range("A5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDashDotDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDashDotDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDashDotDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
MsgBox "Datos borrados, comienze desde el paso 1 ", vbInformation

Else

MsgBox "No se borraron los datos", vbInformation

End If

MsgBox "Macro realizada por Alfredo Saldaña C. (saldanac@) ", vbInformation

End Sub

