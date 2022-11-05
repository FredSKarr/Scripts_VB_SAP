Sub MDC4_putaway()

Dim App, Connection, session As Object
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)
Dim BolVal As Boolean
Dim IntUltimaF, IntUltimaF2, IntUltimaF3, IntContador, IntQty As Integer
Dim StrPN, StrBin, Str313, StrDate1, StrDate2 As String

Application.ScreenUpdating = False

ActiveSheet.Unprotect 'desprotege celdas


'-------------P/N & qty Search --------------------------------------

'-------------------TCODE LT23 ---------------------------

IntUltimaF = (Range("C6").End(xlDown).Offset(1, 0).Row) - 1
Range("C6:C" & IntUltimaF).Copy

session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLT23"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = "020"
session.findById("wnd[0]/usr/btn%_T1_TANUM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/radT1_ALLTA").Select
session.findById("wnd[0]/usr/ctxtLISTV").Text = "/PAT_LT23_LO"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 9
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "PATLT23.txt"
session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "0000"
session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus
session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

Application.CutCopyMode = False

'---------------------Excel File PATLT23.txt -------------------------------------------
Workbooks.OpenText Filename:="C:\PATLT23.txt", Origin:=437, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers _
        :=True
IntUltimaF = (Range("B6").End(xlDown).Offset(1, 0).Row) - 1
Range("D6").Select
Range("D6").FormulaR1C1 = "=RC[-2]&RC[-1]"
Selection.AutoFill Destination:=Range("D6:D" & IntUltimaF)

Windows("IBMPowerMDC4_Put away Tool.xlsm").Activate
Range("F6").FormulaR1C1 = "=VLOOKUP(RC[-3]&RC[-2],PATLT23.txt!R6C4:R300C7,2,0)"
Range("G6").FormulaR1C1 = "=VLOOKUP(RC[-4]&RC[-3],PATLT23.txt!R6C4:R300C7,4,0)"
IntUltimaF = (Range("C6").End(xlDown).Offset(1, 0).Row) - 1
Range("F6:G6").Select
Selection.AutoFill Destination:=Range("F6:G" & IntUltimaF)
Range("F6:G" & IntUltimaF).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("H6").Select
Application.CutCopyMode = False

Windows("PATLT23.txt").Activate
Application.DisplayAlerts = False
ActiveWindow.Close
Application.DisplayAlerts = True

Windows("IBMPowerMDC4_Put away Tool.xlsm").Activate

'-----------------Mov Search ---------------------------------

'-------------------by PN TCODE MB51 ---------------------------
Range("F6:F" & IntUltimaF).Copy
StrDate1 = Range("L3").Text
StrDate2 = Range("M3").Text

session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb51"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtMATNR-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "GD01"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "MDC4"
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = StrDate2
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = StrDate1
session.findById("wnd[0]/usr/radRFLAT_L").Select
session.findById("wnd[0]/usr/ctxtALV_DEF").Text = "/PAT_MB51"
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 9
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "PATMB51.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

Application.CutCopyMode = False

'------------------------Excel File PATMB51.txt ---------------------------

Workbooks.OpenText Filename:="C:\PATMB51.txt", Origin:=437, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 2), Array(5, 1), Array(6, 2), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1)), TrailingMinusNumbers _
        :=True
IntUltimaF = (Range("F6").End(xlDown).Offset(1, 0).Row) - 1
Range("A4").FormulaR1C1 = "=MID(RC[3],4,7)"
Range("B4").FormulaR1C1 = "=MID(RC[2],14,1)"
Columns("C:C").Select
Selection.Insert Shift:=xlToRight
Range("C4").FormulaR1C1 = "=RC[-2]&RC[-1]"
Range("A4:C4").Select
Selection.AutoFill Destination:=Range("A4:C" & IntUltimaF), Type:=xlFillDefault

Windows("IBMPowerMDC4_Put away Tool.xlsm").Activate
IntUltimaF = (Range("F6").End(xlDown).Offset(1, 0).Row) - 1
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5]&RC[-4],PATMB51.txt!R4C3:R1200C6,4,0)"
Range("H6").Select
Selection.AutoFill Destination:=Range("H6:H" & IntUltimaF)
Range("H6:H" & IntUltimaF).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False
Range("H6").Select

Windows("PATMB51.txt").Activate
Application.DisplayAlerts = False
ActiveWindow.Close
Application.DisplayAlerts = True

Windows("IBMPowerMDC4_Put away Tool.xlsm").Activate

'----------------Mov 315 by TOCODE YMMP ------------------------------
IntContador = (Range("I4").End(xlDown).Offset(1, 0).Row)
IntUltimaF = (Range("F6").End(xlDown).Offset(1, 0).Row)
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nymmp"
session.findById("wnd[0]").sendVKey 0


Do While IntContador <> IntUltimaF
    Str313 = Range("H" & IntContador).Text 'asignacion variable mov 313
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nymmp"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtDS-MD-MBLNR").Text = Str313
    session.findById("wnd[0]/usr/txtDS-MD-ZEILE").Text = "1"
    session.findById("wnd[0]/usr/txtDS-MD-MBLNR").caretPosition = 3
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    Range("L" & IntContador).FormulaR1C1 = session.findById("wnd[0]/sbar").Text
    If Range("K36").Text > 0 Then ' validacion de errores en mov 315
        MsgBox "ERRORES FATALES EN LOS MOVIMIENTOS 315, REVISA, CORRIGE y VUELVE A EJECUTAR!!!", vbCritical
        IntContador = IntUltimaF - 1
        BolVal = True
    End If
    IntContador = IntContador + 1 'incremental
    If IntContador = IntUltimaF Then Exit Do 'salida del ciclo
     
Loop
     session.findById("wnd[0]/tbar[0]/btn[15]").press
     
     
'-------------------- Putaway TCODE LT01 ----------------------------------

If BolVal = False Then ' valida sino hay errores en los mov 315
    IntContador = 6
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nlt01"
    session.findById("wnd[0]").sendVKey 0

    Do While IntContador <> 0
        StrPN = Range("F" & IntContador).Text 'asignacion variable PN
        IntQty = Range("G" & IntContador).Text 'asignacion variable qty
        StrBin = Range("E" & IntContador).Text 'asignacion variable BinLoc
        session.findById("wnd[0]/usr/ctxtLTAK-BETYP").Text = "D"
        session.findById("wnd[0]/usr/txtLTAK-BENUM").Text = "reubicared"
        session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").Text = "999"
        session.findById("wnd[0]/usr/ctxtLTAP-MATNR").Text = StrPN
        session.findById("wnd[0]/usr/txtRL03T-ANFME").Text = IntQty
        session.findById("wnd[0]/usr/ctxtLTAP-WERKS").Text = "GD01"
        session.findById("wnd[0]/usr/ctxtLTAP-LGORT").Text = "MDC4"
        session.findById("wnd[0]/usr/ctxtLTAP-LGORT").caretPosition = 4
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").Text = "918"
        session.findById("wnd[0]/usr/txtLTAP-VLPLA").caretPosition = 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").Text = "RAK"
        session.findById("wnd[0]/usr/txtLTAP-NLPLA").Text = StrBin
        session.findById("wnd[0]/usr/txtLTAP-NLPLA").caretPosition = 6
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        Range("M" & IntContador).FormulaR1C1 = session.findById("wnd[0]/sbar").Text
        IntContador = IntContador + 1 'incremental
        If IntContador = IntUltimaF Then Exit Do 'salida del ciclo

    Loop
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    
'----------------- Clean, Save and Keep records on secondary sheet ----------------------
    Range("C6:J" & IntUltimaF).Copy
    Sheets("Registro").Select
    IntUltimaF2 = (Range("B1").End(xlDown).Offset(1, 0).Row)
    Range("B" & IntUltimaF2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    IntUltimaF3 = (Range("B1").End(xlDown).Offset(1, 0).Row) - 1
    Range("A" & IntUltimaF2 & ":A" & IntUltimaF3).FormulaR1C1 = "=TODAY()"
    Range("A" & IntUltimaF2 & ":A" & IntUltimaF3).Select
    Range("A" & IntUltimaF2 & ":A" & IntUltimaF3).Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A" & IntUltimaF3).Select
    Sheets("Datos").Select
    Range("C6:J" & IntUltimaF).ClearContents
    Range("L6:M" & IntUltimaF).ClearContents
    .ClearContents
    Range("C6").Select
    ActiveWorkbook.Save

Else

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'protege celdas

MsgBox "Macro realizada por Alfredo Saldaña C. (saldanac@) ", vbInformation
Application.ScreenUpdating = True

End If

End Sub

