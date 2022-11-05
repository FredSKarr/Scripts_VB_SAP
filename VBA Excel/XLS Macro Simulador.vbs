Sub SimuladorRutas_1()
'
' SimuladorRutas_1 Macro
' Da Formato Al Simulador Bajado de SAP ZSIM
'

    Dim IntUltimaFila As Integer
    Dim IntContador As Integer
    Dim IntContador2 As Integer
    Dim IntContador3 As Integer
    Dim StrVariable As String
    Dim StrVariable2 As String

    'Dim StrIdentificacion As String
    'StrIdentificacion = InputBox("Identificate:", "Password")
    'If StrIdentificacion = "sim2103" Then
    Application.ScreenUpdating = False
    Range("A:B").Delete Shift:=xlToLeft
    Range("A:A").ClearContents
    Range("1:4").ClearContents
    Range("1:3").Insert Shift:=xlDown
    Range("D:D,H:J,L:M,O:O,Q:S,U:U,W:W,Y:AF").Delete Shift:=xlToLeft
    IntUltimaFila = Range("B8").End(xlDown).Offset(1, 0).Row
'-----------Datos de FECHA-----------------------
    Range("G:K").Insert Shift:=xlToRight
    Range("G8").FormulaR1C1 = "=LEFT(RC[-1],2)"
    Range("H8").FormulaR1C1 = "=MID(RC[-2],4,2)"
    Range("I8").FormulaR1C1 = "=RIGHT(RC[-3],4)"
    Range("J8").FormulaR1C1 = "=DATE(RC[-1],RC[-2],RC[-3])"
    Range("G8:J8").AutoFill Destination:=Range("G8" & ":J" & (IntUltimaFila) - 1)
    Columns("G:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Range("F:I").Delete Shift:=xlToLeft
    Range("G8").FormulaR1C1 = "=TODAY()-RC[-1]"
    Range("G8").AutoFill Destination:=Range("G8" & ":G" & (IntUltimaFila) - 1)
    Range("G:G").NumberFormat = "0"
'-----------ENCABEZADOS----------------
    'Range("A7").FormulaR1C1 = "Ruta"
    Range("B7").FormulaR1C1 = "Numéro de Cliente"
    Range("C7").FormulaR1C1 = "Zona de Ventas"
    Range("D7").FormulaR1C1 = "Número de Pedido"
    Range("E7").FormulaR1C1 = "Zona de Transporte"
    Range("F7").FormulaR1C1 = "Fecha de Pedido"
    Range("G7").FormulaR1C1 = "Dias en Simulador"
    Range("H7").FormulaR1C1 = "Importe de Pedido"
    Range("I7").FormulaR1C1 = "Kgs."
    Range("J7").FormulaR1C1 = "m3"
    Range("K7").FormulaR1C1 = "Nivel de Servicio"
    Range("L7").FormulaR1C1 = "m3 Facturable"
    Range("M7").FormulaR1C1 = "Importe Facturable"
    Range("N7").FormulaR1C1 = "Nombre del Cliente"
    Range("O7").FormulaR1C1 = "Población"
    Range("P7").FormulaR1C1 = "Domicilio"
    Range("Q7").FormulaR1C1 = "Colonia"
'--------------------------FORMATO ENCABEZADO --------------------
    Range("A:S").Select
        With Selection
          .Interior.ColorIndex = 2 'celdas Blancas
        End With
    With Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
    Range("A7:Q7").Select
        With Selection
            .Interior.ColorIndex = 46 'celda naranja
            .Interior.Pattern = xlSolid
            .Font.ColorIndex = 2 'letra blanca
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlJustify
            .ReadingOrder = xlContext
    End With
    Range("A:Z").EntireColumn.AutoFit
    Range("E:E,L:L,B:B").ColumnWidth = 11.95
    Range("A7,I7,N7,O7,P7,Q7,R7,T7,Z7").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
'--------------- Inmovilizacion Paneles, Zoom, formatos de celdas -------------------
    Range("F8").Select
        ActiveWindow.FreezePanes = True
        ActiveWindow.Zoom = 85
    Range("A:Z").EntireColumn.AutoFit
    Range("A:M").HorizontalAlignment = xlCenter
    Range("B8" & ":Q" & (IntUltimaFila - 1)).Sort Key1:=Range("E8"), Order1:=xlAscending, Key2:=Range("O8") _
        , Order2:=xlAscending, Key3:=Range("B8"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
        xlSortNormal
    
'---------------------Inserta hoja, cambia nombre ------------------------
    ActiveSheet.Name = "Simulador"
    Sheets.Add
    Sheets("Hoja1").Select
    Sheets("Hoja1").Move After:=Sheets(2)
    Sheets("Hoja1").Select
    ActiveSheet.Name = "BO"
 
 '--------------------Tratamiento de BO ------------------------------------
    StrVariable2 = InputBox("Fecha de BO:", "Tratamiento de BO")
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:= _
        "\\guadalajara\Publico\ALMACEN\Embarque\SIMULADORES\Simulador " & StrVariable2 & ".xls", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Application.Workbooks.Open ("\\guadalajara\Publico\ALMACEN\Embarque\SIMULADORES\BO" & StrVariable2 & ".xls")
    Cells.Select
    Selection.Copy
    Windows("Simulador " & StrVariable2 & ".xls").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("1:5").ClearContents
    Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("B:B,D:D,J:J,L:M,O:O,R:R,T:U").Delete Shift:=xlToLeft
    Range("1:7").Insert Shift:=xlDown
    IntUltimaFila = Range("A8").End(xlDown).Offset(1, 0).Row
'-------------TITULOS BO---------------------------------
    Range("A1:C1").Merge
    Range("A1").FormulaR1C1 = "TRUPER HERRAMIENTAS, S.A. DE C.V"
    Range("A2:C2").Merge
    Range("A2").FormulaR1C1 = "SUCURSAL  GUADALAJARA"
    Range("A3:C3").Merge
    Range("A3").FormulaR1C1 = "ANALISIS B.O. POR GENERAR AL:"
    Range("A4:C4").Merge
    Range("A4").FormulaR1C1 = "=TODAY()"
    Range("A4").NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
    Range("A1:C4").Select
    With Selection
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    Range("A4").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
'---------------IMAGEN LOGO BO---------------------------------
    'Range("J3").Select
    'ActiveSheet.Pictures.Insert("C:\AREA INVENTARIOS\Corel\naranja.jpg").Select
    'Selection.ShapeRange.IncrementLeft -252.75
    'Selection.ShapeRange.IncrementTop -3.75
    'Selection.ShapeRange.ScaleWidth 0.36, msoFalse, msoScaleFromBottomRight
    'Selection.ShapeRange.ScaleHeight 0.36, msoFalse, msoScaleFromTopLeft
    Range("A1").Select
'-----------ENCABEZADOS BO----------------
    Range("A7").FormulaR1C1 = "Numéro Pedido"
    Range("B7").FormulaR1C1 = "Numéro Cliente"
    Range("C7").FormulaR1C1 = "Nombre Cliente"
    Range("D7").FormulaR1C1 = "Codigo"
    Range("E7").FormulaR1C1 = "Descripción"
    Range("F7").FormulaR1C1 = "Promoción Especifico"
    Range("G7").FormulaR1C1 = "Cantidad Solicitada"
    Range("H7").FormulaR1C1 = "Inventario Actual"
    Range("I7").FormulaR1C1 = "Transito a Sucursal"
    Range("J7").FormulaR1C1 = "Existencia Planta 5"
    Range("K7").FormulaR1C1 = "Cant. Back Order"
    Range("L7").FormulaR1C1 = "Importe BO"
    Range("M7").FormulaR1C1 = "Dias Equivalentes"
    Range("N7").FormulaR1C1 = "Estado"
'-----------------FORMATO ENCABEZADO BO --------------------
    Range("A7:N7").Select
        With Selection
            .Interior.ColorIndex = 46 'celda naranja
            .Interior.Pattern = xlSolid
            .Font.ColorIndex = 2 'letra blanca
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlJustify
            .ReadingOrder = xlContext
    End With
    Range("A:N").EntireColumn.AutoFit
    Range("E:E,L:L,B:B").ColumnWidth = 11.95
    Range("C7,D7,E7").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Range("A7" & ":N" & (IntUltimaFila) - 1).Select
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
'--------------- Inmovilizacion Paneles, Zoom, formatos de celdas BO-------------------
    Range("A8").Select
        ActiveWindow.FreezePanes = True
        ActiveWindow.Zoom = 80
    Range("A:N").EntireColumn.AutoFit
    Range("A:B,D:D,F:I,M:N").HorizontalAlignment = xlCenter
    Range("L:L").Style = "Currency"
    Range("C7,E7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .IndentLevel = 0
        .ReadingOrder = xlContext
    End With
    Columns("M:M").ColumnWidth = 13.57
    Rows("7:7").RowHeight = 27.75
    Columns("C:C").ColumnWidth = 32
    Columns("E:E").ColumnWidth = 32
    Columns("I:I").ColumnWidth = 16
    Columns("N:N").ColumnWidth = 16
'--------------Columna Estado de BO (formula Y formato condicional)------------------------------------------
    Range("N8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(Simulador!R1C3:R100C3,BO!RC[-13])>0,""si se trabaja"",""no se trabaja"")"
    Range("N8").AutoFill Destination:=Range("N8" & ":N" & (IntUltimaFila - 1))
    Range("N8" & ":N" & (IntUltimaFila - 1)).Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""si se trabaja"""
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ColorIndex = 2 'letra blanca
    End With
    Selection.FormatConditions(1).Interior.ColorIndex = 3 'Color rojo
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""no se trabaja"""
    With Selection.FormatConditions(2).Font
        .Bold = True
        .Italic = False
        .ColorIndex = xlAutomatic 'Fuente Negra
    End With
    Selection.FormatConditions(2).Interior.ColorIndex = 6 'Color Amarillo
    Range("N8").Select
    Sheets("simulador").Select
    Range("F8").Select

 '------------Modulo Destinatarios---------------------
    Application.Workbooks.Open ("\\guadalajara\Publico\ALMACEN\Bases de Datos\Master Clientes CDR GUA V2013.xls")
    Windows("Simulador " & StrVariable2 & ".xls").Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]="""",VLOOKUP(RC[-1],'[Master Clientes CDR GUA V2013.xls]Base Ctes.'!C2:C15,10,0),RC[-1])"
    Range("C8").Select
    Selection.AutoFill Destination:=Range("C8" & ":C" & IntUltimaFila - 1), Type:=xlFillDefault
    Columns("C:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B7").Select
    Selection.Copy
    Range("C7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
'-------------------------------------------------------------------------
'------------------- Formato de impresion ---------------------------
    ActiveSheet.PageSetup.PrintArea = ("A7" & ":R" & ((IntUltimaFila - 1)))
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
        .PrintErrors = xlPrintErrorsDisplayed
    End With
'--------------------------Nivel de servicio -------------------------------------
    Range(IntUltimaFila & ":" & IntUltimaFila).Insert Shift:=xlDown
    Range("L:L").Insert Shift:=xlToRight
    Range("K7").Copy
    Range("L7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("L8").Select
        With Selection
            .FormulaR1C1 = "=RC[-1]/100"
            .Style = "Percent"
            .AutoFill Destination:=Range("L8" & ":L" & (IntUltimaFila) - 1)
        End With
    IntContador = 7
    IntContador2 = 10  'nivel de servicio minino
    Do While IntContador <> 0
        IntContador = IntContador + 1
        If Range("K" & IntContador).Text < IntContador2 Then
            Range(IntContador & ":" & IntContador).Cut
            Range("A" & (IntUltimaFila + 1)).Insert Shift:=xlDown
            IntUltimaFila = IntUltimaFila - 1
            IntContador = IntContador - 1
        End If
        If IntContador = IntUltimaFila Then Exit Do
        Loop
    Range("L:L").Select
        With Selection
            .Copy
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        End With
    Range("K:K").Delete Shift:=xlToLeft
    Range("A1").Select
'------------------------CLIENTES RECOGE -----------------------------
    IntContador = 7
    IntContador2 = 8
    IntContador3 = 0
    StrVariable = 542325 ' Ferreabastecedora
    Do While IntContador <> 0
        IntContador = IntContador + 1
    If Range("B" & IntContador).Text = StrVariable Then
        Range(IntContador & ":" & IntContador).Cut
        Range("A" & (IntContador2)).Insert Shift:=xlDown
        Range((IntContador2 + 1) & ":" & (IntContador2 + 1)).Insert Shift:=xlDown
        IntContador2 = IntContador2 + 2
        IntUltimaFila = IntUltimaFila + 1
        IntContador3 = IntContador3 + 1
    End If
    If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop
    IntContador = Incontador3 + 7
    StrVariable = 542350 'Arsenio
    Do While IntContador <> 0
        IntContador = IntContador + 1
    If Range("B" & IntContador).Text = StrVariable Then
        Range(IntContador & ":" & IntContador).Cut
        Range("A" & (IntContador2)).Insert Shift:=xlDown
        Range((IntContador2 + 1) & ":" & (IntContador2 + 1)).Insert Shift:=xlDown
        IntContador2 = IntContador2 + 2
        IntUltimaFila = IntUltimaFila + 1
        IntContador3 = IntContador3 + 1
    End If
    If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop
    IntContador = Incontador3 + 7
    StrVariable = 544056 'Pf Mayoristas
    Do While IntContador <> 0
        IntContador = IntContador + 1
    If Range("B" & IntContador).Text = StrVariable Then
        Range(IntContador & ":" & IntContador).Cut
        Range("A" & (IntContador2)).Insert Shift:=xlDown
        Range((IntContador2 + 1) & ":" & (IntContador2 + 1)).Insert Shift:=xlDown
        IntContador2 = IntContador2 + 2
        IntUltimaFila = IntUltimaFila + 1
        IntContador3 = IntContador3 + 1
    End If
    If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop
    IntContador = Incontador3 + 7
    StrVariable = 542362 'Copernico
    Do While IntContador <> 0
        IntContador = IntContador + 1
    If Range("B" & IntContador).Text = StrVariable Then
        Range(IntContador & ":" & IntContador).Cut
        Range("A" & (IntContador2)).Insert Shift:=xlDown
        Range((IntContador2 + 1) & ":" & (IntContador2 + 1)).Insert Shift:=xlDown
        IntContador2 = IntContador2 + 2
        IntUltimaFila = IntUltimaFila + 1
        IntContador3 = IntContador3 + 1
    End If
    If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop
    'IntContador = Incontador3 + 7
    'StrVariable = 102871 'Aceros y Materiales
    'Do While IntContador <> 0
        'IntContador = IntContador + 1
    'If Range("B" & IntContador).Text = StrVariable Then
        'Range(IntContador & ":" & IntContador).Cut
        'Range("A" & (Intcontador2)).Insert Shift:=xlDown
        'Range((Intcontador2 + 1) & ":" & (Intcontador2 + 1)).Insert Shift:=xlDown
        'Intcontador2 = Intcontador2 + 2
        'IntUltimaFila = IntUltimaFila + 1
        'intContador3 = intContador3 + 1
    'End If
    'If IntContador = IntUltimaFila + Intcontador2 Then Exit Do
    'Loop
    IntContador = Incontador3 + 7
    StrVariable = 107318 'Caminante
    Do While IntContador <> 0
        IntContador = IntContador + 1
    If Range("B" & IntContador).Text = StrVariable Then
        Range(IntContador & ":" & IntContador).Cut
        Range("A" & (IntContador2)).Insert Shift:=xlDown
        Range((IntContador2 + 1) & ":" & (IntContador2 + 1)).Insert Shift:=xlDown
        IntContador2 = IntContador2 + 2
        IntUltimaFila = IntUltimaFila + 1
        IntContador3 = IntContador3 + 1
    End If
    If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop
    IntContador = Incontador3 + 7
    StrVariable = 543831 'Tpp
    Do While IntContador <> 0
        IntContador = IntContador + 1
    If Range("B" & IntContador).Text = StrVariable Then
        Range(IntContador & ":" & IntContador).Cut
        Range("A" & (IntContador2)).Insert Shift:=xlDown
        Range((IntContador2 + 1) & ":" & (IntContador2 + 1)).Insert Shift:=xlDown
        IntContador2 = IntContador2 + 2
        IntUltimaFila = IntUltimaFila + 1
        IntContador3 = IntContador3 + 1
    End If
    If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop

'-------------Reune los Pedidos de los CR -----------------------
    IntContador = 8
    IntContador2 = 0
    Do While IntContador <> 0
        IntContador = IntContador + 1
        StrVariable = 542325 'Ferreabastecedora
    If Range("B" & IntContador).Text = StrVariable Then
        Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        IntUltimaFila = IntUltimaFila - 1
        IntContador2 = IntContador2 + 1
    End If
    If IntContador = 40 Then Exit Do
    Loop
    IntContador2 = IntContador2 + 2
    IntContador = 8 + IntContador2
    Do While IntContador <> 0
        IntContador = IntContador + 1
        StrVariable = 542350 'Arsenio
    If Range("B" & IntContador).Text = StrVariable Then
        Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        IntUltimaFila = IntUltimaFila - 1
        IntContador2 = IntContador2 + 1
    End If
    If IntContador = 40 Then Exit Do
    Loop
    IntContador2 = IntContador2 + 2
    IntContador = 8 + IntContador2
    Do While IntContador <> 0
        IntContador = IntContador + 1
        StrVariable = 544056 ' Pf mayoristas
    If Range("B" & IntContador).Text = StrVariable Then
        Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        IntUltimaFila = IntUltimaFila - 1
        IntContador2 = IntContador2 + 1
    End If
    If IntContador = 40 Then Exit Do
    Loop
    IntContador2 = IntContador2 + 2
    IntContador = 8 + IntContador2
    Do While IntContador <> 0
        IntContador = IntContador + 1
        StrVariable = 542362 'Copernico
    If Range("B" & IntContador).Text = StrVariable Then
        Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        IntUltimaFila = IntUltimaFila - 1
        IntContador2 = IntContador2 + 1
    End If
    If IntContador = 40 Then Exit Do
    Loop
    'Intcontador2 = Intcontador2 + 2
    'IntContador = 8 + Intcontador2
    'Do While IntContador <> 0
        'IntContador = IntContador + 1
        'StrVariable = 102871 'Aceros Y materiales
    'If Range("B" & IntContador).Text = StrVariable Then
        'Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        'IntUltimaFila = IntUltimaFila - 1
        'Intcontador2 = Intcontador2 + 1
    'End If
    'If IntContador = 40 Then Exit Do
    'Loop
    IntContador2 = IntContador2 + 2
    IntContador = 8 + IntContador2
    Do While IntContador <> 0
        IntContador = IntContador + 1
        StrVariable = 107318 'Caminante
    If Range("B" & IntContador).Text = StrVariable Then
        Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        IntUltimaFila = IntUltimaFila - 1
        IntContador2 = IntContador2 + 1
    End If
    If IntContador = 40 Then Exit Do
    Loop
    IntContador2 = IntContador2 + 2
    IntContador = 8 + IntContador2
    Do While IntContador <> 0
        IntContador = IntContador + 1
        StrVariable = 543831 'Tpp
    If Range("B" & IntContador).Text = StrVariable Then
        Range((IntContador - 1) & ":" & (IntContador - 1)).EntireRow.Delete
        IntUltimaFila = IntUltimaFila - 1
        IntContador2 = IntContador2 + 1
    End If
    If IntContador = 40 Then Exit Do
    Loop


'----------------------- INSERTA SALTOS DE PAGINA -------------------------
    IntContador = 7 + IntContador3 + IntContador3
    IntContador2 = 0
    StrVariable = Range("E" & IntContador)
    Do While IntContador <> 0
        IntContador = IntContador + 1
        If Range("E" & IntContador).Text <> StrVariable Then
            StrVariable = Range("E" & IntContador)
            Range(IntContador & ":" & IntContador).Insert Shift:=xlDown
            IntContador2 = IntContador2 + 1
        End If
        If IntContador = IntUltimaFila + IntContador2 Then Exit Do
    Loop
'-----------------------ADAPTACIONES -------------------------------
    Range("A:A").Delete Shift:=xlToLeft
    Range("G:G").Insert Shift:=xlToRight
    Range("G7").FormulaR1C1 = "Estatus"
    Range("G7:G500").Select
    With Selection.Font
        .Bold = True
        .ColorIndex = 3
    End With
    Range("1:5").Delete Shift:=xlToUp
    Range("F:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
'----------------------- Tabla Dinamica ----------------------------------
    Range("E2:M200").Select
    ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
        "Simulador!R2C5:R200C13").CreatePivotTable TableDestination:="", TableName _
        :="Tabla dinámica5", DefaultVersion:=xlPivotTableVersion10
    ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
    ActiveSheet.Cells(3, 1).Select
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Estatus")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Fecha de Pedido")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica5").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica5").PivotFields("Importe Facturable"), _
        "Suma de Importe Facturable", xlSum
    Range("C5").Select
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
'---------------------------------Formato Tabla -------------------------------------
    Range("B4:D4").Select
    With Selection
        .Font.Bold = True
        .Font.ColorIndex = 2 'letra blanca
        .Interior.ColorIndex = 46 'celda naranja
        .Interior.Pattern = xlSolid
        .NumberFormat = "ddd dd/mm"
    End With
    Range("A6").Select
    With Selection
        .Interior.ColorIndex = 1
        .Interior.Pattern = xlSolid
        .Font.ColorIndex = 2
        .Font.Bold = True
    End With
    Range("B5").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    Sheets("Hoja2").Select
    Sheets("Hoja2").Name = "Reporte"

'------------------------Instrucciones Finales ---------------------------
    Sheets("simulador").Select
    Range("F8").Select
    Workbooks("BO" & StrVariable2 & ".xls").Close ' Cierra libro BO
    Workbooks("Master Clientes CDR GUA V2013.xls").Close ' Cierra libro Clientes
    Windows("Simulador " & StrVariable2 & ".xls").Activate
    ActiveWorkbook.Save 'Guarda Trabajo
    Application.DisplayAlerts = True ' reactiva advertencias
    Application.ScreenUpdating = True
    MsgBox "Macro realizada por Alfredo Saldaña", vbInformation

    'Else
    
        'MsgBox "PEELAAS!!", vbOKOnly + vbCritical, "Alfredo Saldaña"

    'End If

End Sub

