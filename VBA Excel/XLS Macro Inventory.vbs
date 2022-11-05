Sub Inv_Formula()
'
' Inv_Formula Macro
' Macro grabada el 13/10/2010 por asaldañac
'

'
    Dim IntUltimaF As Integer
    Application.DisplayAlerts = True
    Application.ScreenUpdating = False
    IntUltimaF = (Range("A2").End(xlDown).Offset(1, 0).Row) - 1
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((OR(RC[-1]="""",RC[-1]=""00001-0110"")),VLOOKUP(RC[-3],'[Cedula .xls]Cedula'!C1:C10,12,0),0)"
    Selection.AutoFill Destination:=Range("k2" & ":K" & IntUltimaF)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'------------------------------ SECOND PART ---------------------------------------------------------------


Sub Inv_Captura()
'
' Inv_Captura Macro
' Macro grabada el 13/10/2010 por asaldañac
'

'
    Dim StrVariable As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'Columns("C:C").Select
    'Selection.Cut
    'Range("B1").Select
    'Selection.Insert Shift:=xlToRight
    Columns("L:Q").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:E").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.NumberFormat = "0000000000"
    Selection.ColumnWidth = 10
    Columns("B:B").Select
    Selection.ColumnWidth = 3
    Selection.NumberFormat = "000"
    Columns("C:C").Select
    Selection.NumberFormat = "000000"
    Selection.ColumnWidth = 6
    Columns("D:D").Select
    Selection.NumberFormat = "000000000000"
    Selection.ColumnWidth = 12
    Columns("E:H").Select
    Selection.Delete Shift:=xlToLeft
    StrVariable = InputBox("Fecha de Inventario:", "Captura de Inventario")
    ActiveWorkbook.SaveAs Filename:="C:\AREA INVENTARIOS\Borrar\Inventario " & StrVariable & ".prn", _
        FileFormat:=xlTextPrinter, CreateBackup:=False
    ActiveWorkbook.Close ' Cierra libro
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    End Sub




