Sub Trailer2012()
'
' trailer2012 Macro
' Macro grabada el 16/01/2012 por asaldañac
'
'
    Dim IntContador As Integer
    Dim IntUltimaFila As Integer
    Dim IntContador2 As Double
    Dim StrVariable As String
    Dim StrTransporte As String
    Dim StrAgrupador As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
'---------------- Edicion del reporte del agrupador ------------------
    Sheets("Reporte").Select
    Cells.Select
    With Selection
        .UnMerge
        .EntireColumn.AutoFit
    End With
    Columns("W:AG").Delete Shift:=xlToLeft
    Columns("H:U").Delete Shift:=xlToLeft
    Columns("C:F").Delete Shift:=xlToLeft
    Columns("A:A").Delete Shift:=xlToLeft
    Range("1:4").Delete Shift:=xlToUp
    ActiveSheet.Shapes("Picture 1").Delete
    Columns("B:B").Select
    Selection.Cut
    Range("A1").Insert Shift:=xlToRight
    Columns("C:C").Insert Shift:=xlToRight
    Range("A2:A2000").SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=R[-1]C"
    End With
    Range("A2:A2000").Select
    With Selection
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End With
    Application.CutCopyMode = False
    IntUltimaFila = Range("A1").End(xlDown).Offset(1, 0).Row
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(VALUE(RC[-1])),0,VALUE(RC[-1]))"
    Range("C4").Select
    Selection.AutoFill Destination:=Range("C4:C" & IntUltimaFila - 1)
    Columns("C:C").Select
    Selection.Copy
    Application.CutCopyMode = False
    Range("C4").Select
    Selection.Copy
    Range("E4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("F:F").EntireColumn.AutoFit
    Columns("E:E").Select
    Columns("F:F").EntireColumn.AutoFit
    Columns("D:H").Select
    Range("H1").Activate
    Columns("D:H").EntireColumn.AutoFit
    Range("E4").Select
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Range("F4").Select
    Selection.ClearContents
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(VALUE(RC[1])),0,VALUE(RC[1]))"
    Selection.AutoFill Destination:=Range("D4:D" & (IntUltimaFila) - 1)
    Columns("C:D").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("A:D").Select
    Selection.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],0)"
    IntContador = Range("I1").Text
    Rows("1:" & IntContador).Select
    Selection.Delete Shift:=xlUp
    '-----------------------Pasa los datos al formato ----------------------
    IntUltimaFila = Range("B1").End(xlDown).Offset(1, 0).Row
    Range("A1:A" & (IntUltimaFila) - 1).Select
    Selection.Copy
    Sheets("Formato").Select
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Reporte").Select
    Range("B1:B" & (IntUltimaFila) - 1).Select
    Selection.Copy
    Sheets("Formato").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Reporte").Select
    Range("C1:C" & (IntUltimaFila) - 1).Select
    Selection.Copy
    Sheets("Formato").Select
    Range("D5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    '------------------- Busca descripcion, ubicacion ---------------------
    IntUltimaFila = Range("A5").End(xlDown).Offset(1, 0).Row
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ZMR321!C[-2]:C[8],3,0)"
    Selection.AutoFill Destination:=Range("C5:C" & (IntUltimaFila) - 1)
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ZMR321!C[-4]:C[6],10,0)"
    Selection.AutoFill Destination:=Range("E5:E" & (IntUltimaFila) - 1)
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ZMR321!C[-5]:C[5],11,0)"
    Selection.AutoFill Destination:=Range("F5:F" & (IntUltimaFila) - 1)
    Range("A5:F" & (IntUltimaFila) - 1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Sort Key1:=Range("A5"), Order1:=xlAscending, Key2:=Range("B5") _
        , Order2:=xlAscending, Key3:=Range("E5"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
        xlSortNormal
'---------------------- Saltos de Pagina -------------------------------
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.PageSetup.PrintArea = ("A5:F" & (IntUltimaFila) - 1)
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    IntContador = 5
    IntContador2 = 6
    StrVariable = Range("A" & IntContador)
    Do While IntContador <> 0
        If Range("A" & (IntContador) + 1).Text <> StrVariable Then
            Rows(IntContador2 & ":" & IntContador2).Select
            ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
            StrVariable = Range("A" & IntContador2)
        End If
        IntContador = IntContador2
        StrVariable = Range("A" & IntContador2)
        IntContador2 = IntContador2 + 1
        If IntContador = IntUltimaFila Then Exit Do
    Loop
    '--------------- Asignacion Datos de Cabecera -------------------------------
    StrTransporte = InputBox("Numero de Transporte:", "Datos Cabecera")
    Range("D2") = StrTransporte
    StrAgrupador = InputBox("Numero de Agrupador:", "Datos Cabecera")
    Range("C2") = StrAgrupador
    '------------------Impreme, guarda, elimina ---------------------------------
    'ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    Sheets("ZMR321").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Reporte").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveWorkbook.SaveAs Filename:="C:\Recibo\Trailer " & StrTransporte & "-AG" & StrAgrupador & ".xls"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Macro realizada por Alfredo Saldaña", vbInformation
   
   End Sub     
    