Sub Existencias()
'
' Existencias Macro
' Macro grabada el 12/10/2010 por asaldañac
'

'
    Dim IntUltimaF As Integer
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:G").Select
    Selection.Delete Shift:=xlToLeft
    IntUltimaF = (Range("B6").End(xlDown).Offset(1, 0).Row) - 1
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]=0,0,RC[-2]/RC[-1])"
    Range("E6").Select
    Selection.AutoFill Destination:=Range("E6" & ":E" & IntUltimaF)
    Range("E6:E12111").Select
    Range("B6").Select
End Sub
