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

Public Function Execute(Optional ShtName As String = "", _
    Optional ClearStaleData As Boolean = False, Optional SkipConnect As Boolean = False, _
    Optional KeepAlive As Boolean = False, Optional OutputToSheet As Boolean = False, _
    Optional ByRef MapToModel As Object) As ADODB.Recordset
    If Not SkipConnect Then Connect
    
    recordsAffected = 0
    'Dim Records As ADODB.Recordset
    With Command
        .ActiveConnection = Connection
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    
    Set Records = New Recordset
    Set Records = Command.Execute(recordsAffected)
    
    Application.ODBCTimeout = 60
    'strSql = "SET NOCOUNT ON " & vbCrLf & "BEGIN" & vbCrLf & strSql & "; " & strSql & vbCrLf & "END"
    
    'With Records
    '    .Open Source:=strSql, ActiveConnection:=Connection, Options:=adCmdText
    'End With
    
    
    Output ShtName:=ShtName, MapToModel:=MapToModel, ClearStaleData:=ClearStaleData
    
    If Not KeepAlive Then Disconnect

End Function

Sub Output(Optional ShtName As String = "", _
           Optional ClearStaleData As Boolean, _
           Optional MapToModel As Object)
           ', _
           'Optional ByRef Records As Recordset)
    'Set Records = New ADODB.Recordset
    
    Dim fld As ADODB.field
    
    If ShtName <> "" Then
        Do Until Records.EOF
            xRow = xRow + 1
            xCol = 0
            
            elseCol = 10
            For Each fld In Records.fields
                xCol = xCol + 1
                Data.Cells(1, xCol) = fld.Name
                Data.Cells(xRow, xCol) = fld.Value
               
            Next fld
            Records.MoveNext
            
        Loop
        
    ElseIf Not IsMissing(MapToModel) Then
        Do Until Records.EOF
            'xRow = xRow + 1
            'xCol = 0
            
            elseCol = 10
            For Each fld In Records.fields
                'xCol = xCol + 1
                Debug.Print fld.Name, fld.Value
                'Data.Cells(1, xCol) = fld.Name
                'Data.Cells(xRow, xCol) = fld.value
               
            Next fld
            Debug.Print vbCrLf
            Records.MoveNext
            
        Loop
    
    End If
    'MsgBox (CStr(xRow - 1) & " records fetched.")
    
Finally:
    Disconnect
    
End Sub


Sub OutputRecords(ShtName As String, Optional ClearStaleData As Boolean, Optional ByRef Records As Recordset)
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
        .Open source:=strSql, ActiveConnection:=Connection, Options:=adCmdText
    End With
    
    'F = 1
    
    Do Until BaseQueryModule.Records.EOF
        xRow = xRow + 1
        xCol = 0
        
        elseCol = 10
        For Each fld In Records.fields
            xCol = xCol + 1
            Data.Cells(1, xCol) = fld.Name
            Data.Cells(xRow, xCol) = fld.Value
           
        Next fld
        Records.MoveNext
        
    Loop
    MsgBox (CStr(xRow - 1) & " records fetched.")
    
Finally:
    Disconnect
    
End Sub
