Attribute VB_Name = "Connection"
Public Connection As ADODB.Connection
Public Command As ADODB.Command
Public Records As ADODB.Recordset
Public strConnection As String
Public strSql As String

Function Sql(s As String) As String
    'Append more SQL to the public string and return it
    If Left(s, 1) <> "," Then s = " " & s & vbCrLf
    strSql = strSql & s
    Sql = strSql

End Function

Function Connect() As ADODB.Connection
    'Clear previous SQL strings, (re)open the connection
    On Error GoTo CheckConnectionStatus
    
    'strSql = ""
    Set Connection = New ADODB.Connection
    Set Records = New ADODB.Recordset
    Set Command = New ADODB.Command
    strConnection = "PROVIDER=SQLOLEDB;"
    strConnection = strConnection + "DATA SOURCE=RJOPDB01V;INITIAL CATALOG=PCOM_CONSUMER; "
    strConnection = strConnection + "UID=USER_READ; PWD=@TIMbrasil01"
    
   
CheckConnectionStatus:
    
IsClosed:
    On Error GoTo IsOpen
    Connection.Open (strConnection)
    Set Command.ActiveConnection = Connection
    GoTo Finally
    
IsOpen:
    Connection.Close
    GoTo CheckConnectionStatus
    
Finally:
    
End Function

Function Disconnect() As Boolean
    'Clear previous SQL strings, close the connection
    strSql = ""
    On Error GoTo F
    Connection.Close
    Disconnect = True
    Exit Function
    
F:
    Disconnect = False
    Exit Function

End Function

Sub Fetch(Optional ShtName As String = "SQL")
    
AfterCreateSheet:
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Connect
    
    Call OutputRecords(ShtName)
    
    With Application
        .ScreenUpdating = SU
        .Calculation = xlCalculationAutomatic
    End With

    Disconnect
    Exit Sub
    
End Sub

Public Function Execute(Optional ShtName As String = "SQL", Optional ClearStaleData As Boolean = False, Optional SkipConnect As Boolean = False, Optional KeepAlive As Boolean = False, Optional Output As Boolean = True) As ADODB.Recordset
    If Not SkipConnect Then Connect
    
    'Dim Records As ADODB.Recordset
    'Set Records = Connection.Execute(strSql)
    If Output Then Call OutputRecords(ShtName, ClearStaleData)
    
    If Not KeepAlive Then Disconnect
    
    'Execute = Records

End Function

Sub OutputRecords(ByRef Records As Recordset)
    'Set Records = New ADODB.Recordset
    
    Dim fld As ADODB.field
    
    xRow = 1
    
    Application.ODBCTimeout = 60
    If CreateClustersTable Then
        'Command.CommandText = TableSQL
        'Command.CommandType = adCmdText
        'Set Records.Source = Command
        strSql = "SET NOCOUNT ON " & vbCrLf & "BEGIN" & vbCrLf & TableSql & "; " & strSql & vbCrLf & "END"
        'Debug.Print TableSql

    End If
    
    'Debug.Print Records.Status
    'Set Records = Command.Execute(strSql)
    
    With Records
        .Open Source:=strSql, ActiveConnection:=Connection, Options:=adCmdText
    End With
    
    'F = 1
    
    Do Until BaseQueryModule.Records.EOF
        xRow = xRow + 1
        xCol = 0
        
        elseCol = 10
        For Each fld In Records.fields
            xCol = xCol + 1
            Data.Cells(1, xCol) = fld.name
            Data.Cells(xRow, xCol) = fld.value
           
        Next fld
        Records.MoveNext
        
    Loop
    MsgBox (CStr(xRow - 1) & " records fetched.")
    
Finally:
    Disconnect
    
End Sub
