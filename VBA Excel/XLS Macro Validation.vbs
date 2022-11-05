Sub Validacion()

If Range("IU65536") > 24 Then
    ActiveSheet.Name = "Form Ctes."
    Sheets.Add
    Sheets("Base Ctes.").Visible = False
    Sheets("Form Ctes.").Visible = False
    
    MsgBox "Archivo DESACTIVADO informa a Gerencia", vbInformation

End If
If Range("IU65536") > 16 Then
        MsgBox "Debes ACTUALIZAR este Archivo de lo contrario se DESACTIVARA", vbInformation

End If

End Sub