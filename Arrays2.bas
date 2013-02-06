Attribute VB_Name = "Arrays2"
Public Function ReturnArray(Arr, Optional ByRef Application_Caller As Object)
    If IsMissing(Application_Caller) Then
        CallerRows = UBound(Arr)
        CallerCols = 1
    Else
        CallerRows = Application_Caller.Rows.Count
        CallerCols = Application_Caller.Columns.Count
    End If

    ReDim Result(1 To CallerRows, 1 To CallerCols)
    For RowNdx = 1 To CallerRows
        For ColNdx = 1 To CallerCols
            N = N + 1
            Result(RowNdx, ColNdx) = GetItem(Arr, N)
        Next ColNdx
     Next RowNdx
    
    ReturnArray = Result

End Function

Function ExcludeEmpty(ByRef Arr As Variant)
    Dim Results As Variant
    ReDim Results(1 To 1)
    
    For x = LBound(Arr) To UBound(Arr)
        If Arr(x) <> vbEmpty Then Results(UBound(Results)) = Arr(x)
    Next
    
    ExcludeEmpty = Results
    
End Function

Function PresentInArray(InputArray, Value)
    For i = LBound(InputArray) To UBound(InputArray)
        If InputArray(i) = Value Then
            PresentInArray = True
            Exit Function
        End If
    Next i
    
    PresentInArray = False
End Function

Function AppendToArrayUniquely(InputArray, Value) As Boolean
    Dim bool_ As Boolean
    bool_ = False
    
    If Not PresentInArray(InputArray, Value) Then
        bool_ = InsertElementIntoArray(InputArray, UBound(InputArray) + 1, Value)
    End If
    
    AppendToArrayUniquely = bool_
        
End Function

Function GetItem(Arr, Index, Optional Default As String = "")
    On Error GoTo ErrHandler
    GetItem = Arr(Index)
    Exit Function
    
ErrHandler:
    GetItem = Default

End Function
