Sub Actualizacion_Ubicaciones()
'
' Actualizacion_Ubicaciones Macro
' Actualiza el query con las ubicaciones
'

'
    Dim StrIdentificacion As String
    StrIdentificacion = InputBox("Identificate:", "Password")

    If StrIdentificacion = "au2182" Then
    
    Dim IntUltimaF As Integer
    Dim IntUltimaQ As Double
    Dim IntUltimaC As Integer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    '------------- Creacion respaldo del query ------------------------------
    Windows("Master Codigos CDR GUA V2013.xls").Activate
    ActiveWorkbook.SaveAs Filename:= _
        "C:\luis\Respaldo Master Codigos CDR GUA V2013.xls", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    ActiveWorkbook.SaveAs Filename:= _
        "\\guadalajara\Publico\ALMACEN\Bases de Datos\Master Codigos CDR GUA V2013.xls", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    '-------------     Edicion del archivo de Texto --------------------------------
    Windows("PIKEO.xls").Activate
    Worksheets("PIKEO").Select
    IntUltimaF = (Range("B1").End(xlDown).Offset(1, 0).Row) - 1
    Range("A1").AutoFill Destination:=Range("A1" & ":A" & IntUltimaF)
    Range("A:A").ColumnWidth = 4
    Range("B:B").ColumnWidth = 6
    Range("C:C").ColumnWidth = 3
    Range("D:D").ColumnWidth = 4
    ActiveWorkbook.SaveAs Filename:= _
        "C:\luis\PIKEO.prn", FileFormat:= _
        xlTextPrinter, CreateBackup:=False
    ActiveWorkbook.SaveAs Filename:= _
        "C:\luis\PIKEO.xls", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Range("E1").FormulaR1C1 = "=VLOOKUP(RC[-2],Pasillos!R1C1:R36C2,2,0)"
    Range("E1").AutoFill Destination:=Range("E1" & ":E" & IntUltimaF)
    Range("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    '------------------- Actualizacion QUERY ------------------------------------
    Windows("Master Codigos CDR GUA V2013.xls").Activate
    Sheets("Codigos").Select
    IntUltimaQ = (Range("B4").End(xlDown).Offset(1, 0).Row) - 1
    'Range("D:F").EntireColumn.Hidden = False
    'Range("F:F").Insert shift:=xlToRight
    Range("R5").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-16],[PIKEO.xls]PIKEO!C2:C5,4,0)),RC[2],VLOOKUP(RC[-16],[PIKEO.xls]PIKEO!C2:C5,4,0))"
    Range("S5").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-17],[PIKEO.xls]PIKEO!C2:C5,3,0)),RC[2],VLOOKUP(RC[-17],[PIKEO.xls]PIKEO!C2:C5,3,0))"
    Range("R5:F5").AutoFill Destination:=Range("E5" & ":F" & IntUltimaQ)
    Range("E5" & ":F" & IntUltimaQ).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select
    Workbooks("PIKEO.xls").Close
    Application.DisplayAlerts = True
    Windows("Master Codigos CDR GUA V2013.xls").Activate
    Range("R5" & ":S" & IntUltimaQ).ClearContents
    'Range("F:F").Delete shift:=xlToLeft
    'Range("E:E").EntireColumn.Hidden = True
    Range("B4" & ":Q" & IntUltimaQ).Sort Key1:=Range("D5"), Order1:=xlAscending, Key2:=Range("E5") _
        , Order2:=xlAscending, Key3:=Range("B5"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
        xlSortNormal
    '--------------------- Actualizacion Consolidado -----------------------------
    Sheets("Consolidado Ubicaciones").Select
    IntUltimaC = (Range("A4").End(xlDown).Offset(1, 0).Row) - 1
    Range("G4").FormulaR1C1 = _
        "=IF(RC[-6]=""SIS"",VLOOKUP(RC[-5],Codigos!C[-6]:C,3,0),RC[-3])"
    Range("H4").FormulaR1C1 = _
        "=IF(RC[-7]=""SIS"",VLOOKUP(RC[-6],Codigos!C[-7]:C[-1],4,0),RC[-3])"
    Range("G4:H4").AutoFill Destination:=Range("H4" & ":I" & IntUltimaC)
    Range("G4" & ":H" & IntUltimaC).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("G:H").ClearContents
    'Range("G3").FormulaR1C1 = "1"
    'Range("H3").FormulaR1C1 = "2"
    Range("A4" & ":F" & IntUltimaC).Sort Key1:=Range("F4"), Order1:=xlAscending, Key2:=Range("E4") _
        , Order2:=xlAscending, Key3:=Range("B4"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
        xlSortNormal
    ActiveWorkbook.Save
    Range("A3").Select
    Application.ScreenUpdating = True
    Else
    
        MsgBox "PEELAAS!!", vbOKOnly + vbCritical, "Alfredo Saldaña"

    End If


    
End Sub


