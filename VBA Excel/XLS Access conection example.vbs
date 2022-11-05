Option Explicit

Const PATH As String = "C:\Users\aledg\OneDrive\Documentos\Curso VBA\Cap. 27\database2.accdb"
Const PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
Const PW As String = vbNullString


Public Function openDBConnection() As Boolean
    Dim cnn As ADODB.Connection: Set cnn = New ADODB.Connection
    
    On Error GoTo connection_error
    With cnn
        .ConnectionString = "Provider=" & PROVIDER & ";Data Source=" & PATH & ";Jet OLEDB:Database Password=" & PW & ";"
        .Open
        Debug.Print .State
        .Close
        Debug.Print .State
    End With
    
    Set cnn = Nothing
    openDBConnection = True
    Exit Function
    
connection_error:
    Dim error_msg As String, i As Integer
    For i = 0 To cnn.Errors.Count - 1
        error_msg = error_msg & IIf(error_msg = vbNullString, vbNullString, vbNewLine) & cnn.Errors(i)
    Next i
    
    openDBConnection = False
    Msg "¡No se ha podido realizar la conexión a la base de datos!" & vbNewLine & vbNewLine & error_msg, 2, "Database connection failed"
    Set cnn = Nothing
End Function
