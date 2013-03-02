Attribute VB_Name = "ORM"
Function CharField(db_column As String) As Object
    Dim c As ORMCharField
    Set c = New ORMCharField
    c.db_column = db_column
    Set CharField = c
End Function

Sub TestCharField()
    Dim c As ORMCharField
    Set c = CharField("FOO")
    If c.db_column <> "FOO" Then MsgBox ("TestCharField failed!" & vbCrLf & _
                                         c.db_column & " != " & "FOO")
    
End Sub

Function filter(ParamArray ff() As Variant) As String
    filter = vbCrLf & "WHERE"
    Dim x As Long, u As Long, pass As Long
    
    'FIXME: Ignore filters that don't form a pair
    u = UBound(ff) - ((UBound(ff) - LBound(ff)) Mod 2)
    pass = 0
    For x = LBound(ff) To u Step 2
        If pass > 0 Then filter = filter & " AND "
        filter = filter & vbCrLf & "    "
        filter = filter & ParseLookup(ff(x), ff(x + 1))
        pass = pass + 1
        
    Next
        
End Function

Function ParseBaseModel() As String
    With Application.VBE.ActiveVBProject.VBComponents("BaseModel").CodeModule
        ParseBaseModel = .Lines(1, .CountOfLines)
    End With

End Function

Function ParseLookup(ByVal lookup_ As String, ByVal condition As Variant) As String
    Dim lookup_type As String, s As String
    Dim split_pos As Long
    split_pos = InStr(lookup_, "__")
    If split_pos = 0 Or split_pos = Len(lookup_) Then
        lookup_type = "exact"
    Else
        lookup_type = Split(lookup_, "__")(1)
    End If
    lookup_column = Split(lookup_, "__")(0)
    
    s = lookup_column & " "
    'On Error GoTo ConditionIsString
    If IsArray(condition) Then
        condition = SanitizedJoin(condition)
    Else
        condition = SanitizedJoin(¨(condition))
    End If
    
    Select Case lookup_type
    
        Case "exact"
            s = s & "= " & condition
        
        Case "iexact"
            s = s & "= UPPER(" & condition & ")"
        
        Case "in"
            s = s & "IN (" & condition & ")"
        
        Case "startswith"
            s = s & "LIKE " & condition & "&"
            
    End Select
    
    
    ParseLookup = s
    

End Function


Function SanitizedJoin(arr As Variant) As String
    Dim x As Long, pass As Long
    pass = 0
    For x = LBound(arr) To UBound(arr)
        If pass > 0 Then s = s & ", "
        s = s & "'" & arr(x) & "'"
        pass = pass + 1
    Next
    
    SanitizedJoin = s
End Function
