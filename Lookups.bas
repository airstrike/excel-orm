Attribute VB_Name = "Lookups"
Public Function AndreLookup(Field As String, Grouped As Boolean, Sorted As Boolean, _
                         ParamArray Lookups() As Variant) As Variant

    'Application.Volatile
    
    Dim FieldPos As Long, L As Long, NonBlankFilters As Long, ResultSize As Long
    Dim Results() As Variant, Matches() As Variant, _
        LookupFields() As Variant, LookupValues() As Variant, _
        LookupPos() As Variant
    
    Dim Append As Boolean
    ReDim Results(0 To 0) As Variant
    ReDim Matches(0 To 0) As Variant
    ReDim LookupFields(0 To 0) As Variant
    ReDim LookupValues(0 To 0) As Variant
    ReDim LookupPos(0 To 0) As Variant
    Dim LastValue As Variant
    Dim FieldCount As Long
    Dim x As Long
    
    'On Error GoTo ErrHandler
    FieldPos = GetFieldPos(Field)
        
    L = UBound(Lookups) - LBound(Lookups) + 1
    
    If FieldPos = 0 Then AndreLookup = Results
    
    NonBlankFilters = L
    FieldCount = 0
    ResultsSize = 0
    MaxResultsSize = Application.Caller.Rows.Count * Application.Caller.Columns.Count
    
    If L > 0 Then
        For i = LBound(Lookups) To UBound(Lookups) Step 2
            If ((Lookups(i) <> "") And (Lookups(i + 1) <> "")) Then
                FieldCount = FieldCount + 1
            Else
                NonBlankFilters = NonBlankFilters - 2
            End If
        Next
        
        If NonBlankFilters > 0 Then FieldCount = FieldCount - 1

        ReDim LookupFields(0 To FieldCount)
        ReDim LookupValues(0 To FieldCount)
        ReDim LookupPos(0 To FieldCount)
        
        If NonBlankFilters <= 0 Then GoTo StartReturn
        
        x = 0
        For i = LBound(Lookups) To UBound(Lookups) Step 2
            If ((Lookups(i) <> "") And (Lookups(i + 1) <> "")) Then
                LookupFields(x) = CStr(Lookups(i))
                LookupValues(x) = CStr(Lookups(i + 1))
                LookupPos(x) = GetFieldPos(Lookups(i))
                x = x + 1
            End If

        Next i
        
    Else
        FieldCount = 0
    
    End If
    
StartReturn:
    
    With ActiveWorkbook.Sheets("Base")
        lastrow = .UsedRange.Rows.Count
        For xRow = 2 To lastrow Step 1
            Append = True

            If L = 0 Or NonBlankFilters = 0 Then
                InsertedValue = .Cells(xRow, FieldPos).Value
            Else

                For xField = LBound(LookupFields) To UBound(LookupFields)
                    Rowvalue = .Cells(xRow, LookupPos(xField)).Value
                    If LookupValues(xField) <> "" And CStr(Rowvalue) <> LookupValues(xField) Then
                        Append = False
                        GoTo SkipAppending
                    End If
                Next xField
                InsertedValue = .Cells(xRow, FieldPos).Value
            
            End If
            
            If Append = True And LastValue <> InsertedValue Then
                LastValue = InsertedValue
                Inserted = False
                Inserted = AppendToArrayUniquely(Matches, InsertedValue)
                
                If Inserted = True Then
                    ResultsSize = ResultsSize + 1
                    If ResultsSize >= MaxResultsSize Then GoTo ReturnResults
                End If

            End If

SkipAppending:
        Next xRow
        
    End With
    
ReturnResults:
    If Sorted Then
        Call QSortInPlace(Matches)
    End If
    
    On Error GoTo SimpleReturn
    AndreLookup = ReturnArray(Matches, Application.Caller)
    Exit Function
    
SimpleReturn:
    AndreLookup = Matches
    Exit Function
    
ErrHandler:
    AndreLookup = 0
    Exit Function
        

End Function

Public Function UniqueLookup(Field As String, Optional Sorted As Boolean = False) As Variant
    Dim FieldPos As Long, xRow As Long, ReturnRows As Long
    Dim Results() As Variant
    ReDim Results(0 To 0) As Variant
    
    FieldPos = GetFieldPos(Field)
    
    ResultsSize = 0
    'MaxResultsSize = 10000
    MaxResultsSize = Application.Caller.Rows.Count * Application.Caller.Columns.Count
    
    With ActiveWorkbook.Sheets("Base")
        lastrow = .UsedRange.Rows.Count
        
        For xRow = 2 To lastrow Step 1
            InsertedValue = .Cells(xRow, FieldPos).Value
            If LastValue <> InsertedValue Then
                LastValue = InsertedValue
                Inserted = False
                Inserted = AppendToArrayUniquely(Results, InsertedValue)
                
                If Inserted = True Then
                    ResultsSize = ResultsSize + 1
                    If ResultsSize >= MaxResultsSize Then GoTo ReturnResults
                End If
            End If
                        
        Next xRow
    
    End With
    
ReturnResults:
    
    If Sorted Then
        Call QSortInPlace(Results)
    End If
    
    On Error GoTo SimpleReturn
    UniqueLookup = ReturnArray(Results, Application.Caller)
    Exit Function
    
SimpleReturn:
    UniqueLookup = Results
    Exit Function
                                                  
End Function

Public Function GetFieldPos(ByVal Field As String)
    On Error GoTo ErrHandler
    With Application.WorksheetFunction
        GetFieldPos = .Match(Field, ActiveWorkbook.Sheets("BASE").Range("1:1"), 0)
        Exit Function
    End With
    
ErrHandler:
    GetFieldPos = 0
    On Error GoTo 0
    Exit Function
    
End Function
