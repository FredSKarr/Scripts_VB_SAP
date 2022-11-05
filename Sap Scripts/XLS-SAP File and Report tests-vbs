Sub MovAudit()
Dim App, Connection, session As Object
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)
Dim StrDateLow, StrDateUp As String 'variable rango de fechas
Dim IntUltimaF As Integer 'variable para la ultima fila con datos

'Inputbox para preguntar el rango de fechas
StrDateLow = InputBox("Fecha Inicio - formato MMDDAA:", "Rango de Fechas")
StrDateUp = InputBox("Fecha Hasta - formato MMDDAA:", "Rango de Fechas")

'launch a transaction
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nLX02" 'TCODE LX02 (BinLoc)
session.findById("wnd[0]").sendVKey 0

' transaction parameters and commands
session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = "020" 'Warehouse number
session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").Text = "RAK" 'Storage Type
session.findById("wnd[0]/usr/ctxtS1_LGPLA-LOW").Text = "40*" 'Storage bin
session.findById("wnd[0]/usr/ctxtP_VARI").Text = "STDLIST" 'Layout
session.findById("wnd[0]/tbar[1]/btn[8]").press 'ejecute (f8)
session.findById("wnd[0]/tbar[1]/btn[9]").press 'send to file (f9)
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\" 'File path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ScripVBASAP_LX02.txt" ' File export Name
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3 'back
session.findById("wnd[0]").sendVKey 3 'back

'launch a transaction
session.findById("wnd[0]/tbar[0]/okcd").Text = "lt22" 'TCODE LT22 (TO´s)
session.findById("wnd[0]").sendVKey 0

'transaction parameters
session.findById("wnd[0]/usr/radT3_QUITA").Select
session.findById("wnd[0]/usr/chkT3_SENAC").Selected = True
session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = "020" 'warehose num
session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = "RAK" 'Storage Type
session.findById("wnd[0]/usr/txtT3_LGPLA-LOW").Text = "40*" 'Storage bin
session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = StrDateLow 'rango de fechas
session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = StrDateUp 'rango de fechas
session.findById("wnd[0]/usr/ctxtLISTV").Text = "/ALFRED2" 'layout
session.findById("wnd[0]/usr/ctxtLISTV").SetFocus
session.findById("wnd[0]/usr/ctxtLISTV").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[9]").press 'send to file (f9)
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\" 'File path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ScripVBASAP_LT22.txt" ' File export Name
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3 'back
session.findById("wnd[0]").sendVKey 3 'back

'-----------------------------apertura de archivos -----------------------------------------------

'LT22
ChDir "C:\"
    Workbooks.OpenText Filename:="C:\ScripVBASAP_LT22.txt", Origin:=437, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1)), TrailingMinusNumbers:=True
 'LX02
  Workbooks.OpenText Filename:="C:\ScripVBASAP_LX02.txt", Origin:=437, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1)), _
        TrailingMinusNumbers:=True
    
Windows("ScripVBASAP_LT22.txt").Activate

    

    
    '-------------Edición, Formato y encabezado ---------------------------------
    Rows("5:5").Delete Shift:=xlUp
    Columns("D:E").Delete Shift:=xlToLeft
    Range("A4").FormulaR1C1 = "Random Numb"
    Range("Q4").FormulaR1C1 = "Act Source Exist"
    Range("R4").FormulaR1C1 = "Act Dest Exist"
    Range("S4").FormulaR1C1 = "Count Souce"
    Range("T4").FormulaR1C1 = "Count Dest"
    Columns("B:T").EntireColumn.AutoFit
    IntUltimaF = (Range("B4").End(xlDown).Offset(1, 0).Row) - 1 'da valor a la variable
    Range("B4" & ":T" & IntUltimaF).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    '_--------------------Formula matrix --------------------------------------
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "=RANDBETWEEN(1,2)"
    Range("A5").Select
    Selection.AutoFill Destination:=Range("A5" & ":A" & IntUltimaF)
    Range("A5" & ":A" & IntUltimaF).Select
    Selection.Copy
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Q5").Select
    Selection.FormulaArray = _
        "=INDEX(ScripVBASAP_LX02.txt!R8C9:R2000C9,MATCH(RC[-9]&RC[-7],ScripVBASAP_LX02.txt!R8C2:R2000C2&ScripVBASAP_LX02.txt!R8C8:R2000C8,0),0)"
    Range("R5").Select
    Selection.FormulaArray = _
        "=INDEX(ScripVBASAP_LX02.txt!R8C9:R2000C9,MATCH(RC[-10]&RC[-6],ScripVBASAP_LX02.txt!R8C2:R2000C2&ScripVBASAP_LX02.txt!R8C8:R2000C8,0),0)"
    Range("Q5:R5").Select
    Selection.AutoFill Destination:=Range("Q5" & ":R" & IntUltimaF)
    Range("Q5" & ":R" & IntUltimaF).Select
    Selection.Copy
    Range("Q5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Replace What:="#N/A", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.CutCopyMode = False

 '--------------- Filtrado e impresion ---------------------------------------
    Range("A5" & ":T" & IntUltimaF).Select
    ActiveWorkbook.Worksheets("ScripVBASAP_LT22").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ScripVBASAP_LT22").Sort.SortFields.Add Key:=Range( _
        "J5" & ":J" & IntUltimaF), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ScripVBASAP_LT22").Sort
        .SetRange Range("A4" & ":T" & IntUltimaF)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A4").Select
    Selection.AutoFilter
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=1, Criteria1:="2"
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=1
    ActiveSheet.PageSetup.PrintArea = "B4" & ":R" & IntUltimaF
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=1, Criteria1:="2"
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=10, Criteria1:="=40*", _
        Operator:=xlAnd
    Columns("B:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:R").Select
    Selection.EntireColumn.Hidden = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Columns("A:T").Select
    Range("T1").Activate
    Selection.EntireColumn.Hidden = False
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=1
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=10
    ActiveWorkbook.Worksheets("ScripVBASAP_LT22").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ScripVBASAP_LT22").AutoFilter.Sort.SortFields.Add _
        Key:=Range("L5" & ":L" & IntUltimaF), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ScripVBASAP_LT22").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=1, Criteria1:="2"
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=12, Criteria1:="=40*", _
        Operator:=xlAnd
    Columns("B:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("N:Q").Select
    Selection.EntireColumn.Hidden = True
    Columns("M:M").Select
    Selection.EntireColumn.Hidden = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Columns("A:S").Select
    Selection.EntireColumn.Hidden = False
    Range("B9").Select
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=11
    Windows("ScripVBASAP_LX02.txt").Activate
    ActiveWindow.Close
    Windows("ScripVBASAP_LT22.txt").Activate
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=1
    ActiveSheet.Range("A4" & ":T" & IntUltimaF).AutoFilter Field:=11
    ActiveWorkbook.SaveAs Filename:="C:\Mov MDC4 Audit" & StrDateLow & ".xls", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    MsgBox "Todos los archivos fueron guardados en directorio raíz C:", vbInformation
    MsgBox "Macro realizada por Alfredo Saldaña C. (saldanac@) ", vbInformation
End Sub



