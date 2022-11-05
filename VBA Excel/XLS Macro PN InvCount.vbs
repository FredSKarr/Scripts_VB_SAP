Sub Macro1()
'
' Macro1 Macro
' Macro grabada el 08/03/2013 por asaldañac
'

'
    Dim IntUltimaFila As Double
    Dim IntContador As Double
    'Application.ScreenUpdating = False
    IntUltimaFila = Range("B5").End(xlDown).Offset(1, 0).Row
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    IntContador = 6
    Do While IntContador <> 0
        Range("C" & (IntContador) - 1).Copy
        Range("A" & IntContador).Select
        ActiveSheet.Paste
        Range("F" & (IntContador) - 1).Copy
        Range("B" & IntContador).Select
        ActiveSheet.Paste
        If IntContador < (IntUltimaFila) - 1 Then
            Range("A" & IntContador).AutoFill Destination:=Range("A" & IntContador & ":A" & (IntUltimaFila) - 1)
            Range("B" & IntContador).AutoFill Destination:=Range("B" & IntContador & ":B" & (IntUltimaFila) - 1)
        End If
        IntContador = IntUltimaFila + 2
        IntUltimaFila = Range("C" & (IntContador) - 1).End(xlDown).Offset(1, 0).Row
        If IntContador = 65537 Then Exit Do
    Loop
    Columns("D:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:O").Select
    Selection.Delete Shift:=xlToLeft
End Sub

'---------------------------------------------SECOND PART ---------------------------------------------------

   Sub Conteos2()
    Dim IntUltimaFila As Double
    Dim IntContador As Double
    Dim IntContador2 As Double
    Dim StrVariable As String
    'Application.ScreenUpdating = False
    IntUltimaFila = Range("A1").End(xlDown).Offset(1, 0).Row
    IntContador = 1
    IntContador2 = 2
    StrVariable = Range("A" & IntContador + 1)
    Do While IntContador <> 0
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("D" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("E" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("F" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("G" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("H" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("I" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("J" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("J" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("K" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("L" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("M" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("N" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("O" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("P" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("Q" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("R" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("S" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("T" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("U" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("V" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("W" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("X" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("Y" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("Z" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AA" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AB" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AC" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AD" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AE" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AF" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AG" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AH" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AI" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AJ" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AK" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AL" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AM" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AN" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AO" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AP" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AQ" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AR" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AS" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AT" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AU" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AV" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AW" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AX" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AY" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("AZ" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("BA" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            Range("C" & IntContador2).Cut
            Range("BB" & IntContador).Select
            ActiveSheet.Paste
            IntContador2 = IntContador2 + 1
            StrVariable = Range("A" & IntContador2)
        End If

        IntContador = IntContador2
        IntContador2 = IntContador2 + 1
        StrVariable = Range("A" & IntContador2)
        If IntContador = IntUltimaFila Then Exit Do
    Loop

   End Sub
   
'----------------------------------------------------------THIRD PART ------------------------------------------

   Sub Conteos3()
    Dim IntUltimaFila As Double
    Dim IntContador As Double
    Dim IntContador2 As Double
    Dim StrVariable As String
    Dim StrVariable2 As String
    'Application.ScreenUpdating = False
    IntUltimaFila = Range("A1").End(xlDown).Offset(1, 0).Row
    IntContador = 1
    IntContador2 = 2
    StrVariable = Range("A" & IntContador + 1)
    StrVariable2 = Range("B" & IntContador + 1)
    Do While IntContador <> 0
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("F" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("G" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("H" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("I" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("J" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("K" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("L" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("M" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("N" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("O" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("P" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("Q" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("R" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("S" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("T" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("U" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("V" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("W" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        If Range("A" & IntContador).Text = StrVariable Then
            If Range("B" & IntContador).Text = StrVariable2 Then
                Range("C" & IntContador2).Cut
                Range("X" & IntContador).Select
                ActiveSheet.Paste
                Range("D" & IntContador2).Cut
                Range("Y" & IntContador).Select
                ActiveSheet.Paste
                Range("E" & IntContador2).Cut
                Range("Z" & IntContador).Select
                ActiveSheet.Paste
                IntContador2 = IntContador2 + 1
                StrVariable = Range("A" & IntContador2)
                StrVariable2 = Range("B" & IntContador2)
            End If
        End If
        IntContador = IntContador2
        IntContador2 = IntContador2 + 1
        StrVariable = Range("A" & IntContador2)
        StrVariable2 = Range("B" & IntContador2)
        If IntContador = IntUltimaFila Then Exit Do
    Loop

   End Sub
   