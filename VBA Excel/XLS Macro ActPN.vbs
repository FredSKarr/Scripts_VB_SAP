Sub Actualiza_codigos()
'
' Actualiza_codigos Macro
' Actualiza la BD de Codigos en Base a SAP ZMR321, ZVCAT Y ZQVOLUM
'

'
    Dim DbUltimaFila As Double
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    
    '---------------Borrado Inicial ----------------------------
    Sheets("Codigos").Select
    DbUltimaFila = (Range("B5").End(xlDown).Offset(1, 0).Row) - 1
    Range("A5:P" & DbUltimaFila).ClearContents
    '-------------------Edicion ZMR321 ----------------------
    Sheets("ZMR321").Select
    Range("A1:O50000").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    DbUltimaFila = (Range("C1").End(xlDown).Offset(1, 0).Row) - 1
    Range("A1:A" & DbUltimaFila).Copy
    Sheets("Codigos").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("ZMR321").Select
    Range("C1:C" & DbUltimaFila).Copy
    Sheets("Codigos").Select
    Range("C5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("ZMR321").Select
    Range("J1:K" & DbUltimaFila).Copy
    Sheets("Codigos").Select
    Range("D5:E5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("ZMR321").Select
    Range("H1:H" & DbUltimaFila).Copy
    Sheets("Codigos").Select
    Range("Q5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    '--------------------Rellenado de datos ---------------------
    Sheets("Codigos").Select
    DbUltimaFila = (Range("B5").End(xlDown).Offset(1, 0).Row) - 1
    Range("F5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-4],ZQVOLUM!C2:C16,5,0)),"""",VLOOKUP(RC[-4],ZQVOLUM!C2:C16,5,0))"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-5],ZQVOLUM!C2:C16,6,0)),"""",VLOOKUP(RC[-5],ZQVOLUM!C2:C16,6,0))"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-6],ZQVOLUM!C2:C16,7,0)),"""",VLOOKUP(RC[-6],ZQVOLUM!C2:C16,7,0))"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-7],ZQVOLUM!C2:C16,10,0)),"""",VLOOKUP(RC[-7],ZQVOLUM!C2:C16,10,0))"
    Range("J5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-8],ZQVOLUM!C2:C16,11,0)),"""",VLOOKUP(RC[-8],ZQVOLUM!C2:C16,11,0))"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-9],ZQVOLUM!C2:C16,12,0)),"""",VLOOKUP(RC[-9],ZQVOLUM!C2:C16,12,0))"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-10],ZQVOLUM!C2:C16,3,0)),"""",VLOOKUP(RC[-10],ZQVOLUM!C2:C16,3,0))"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-11],ZVCAT!C1:C12,9,0)),"""",VLOOKUP(RC[-11],ZVCAT!C1:C12,9,0))"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-12],ZVCAT!C1:C12,10,0)),"""",VLOOKUP(RC[-12],ZVCAT!C1:C12,10,0))"
    Range("O5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-13],ZVCAT!C1:C12,10,0)),"""",VLOOKUP(RC[-13],ZVCAT!C1:C12,10,0)*1.16)"
    Range("P5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-14],ZVCAT!C1:C12,5,0)),"""",VLOOKUP(RC[-14],ZVCAT!C1:C12,5,0))"
    Range("F5:P5").Copy
    Range("F5:P" & DbUltimaFila).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("F5:P" & DbUltimaFila).Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    '----------------------Actualizacion Consolidado Ubicaciones--------
   
    
    '-------------------------Borrado de Datos ------------------------
    Sheets("ZMR321").Select
    Range("A:O").ClearContents
    Sheets("ZVCAT").Select
    Range("A:M").ClearContents
    Sheets("ZQVOLUM").Select
    Range("A:S").ClearContents
    
    '------------------------Instrucciones Finales ---------------------
    ActiveWorkbook.Save 'Guarda Trabajo
    Application.DisplayAlerts = True ' reactiva advertencias
    Application.ScreenUpdating = True
    MsgBox "Macro realizada por Alfredo Saldaña", vbInformation


End Sub
 