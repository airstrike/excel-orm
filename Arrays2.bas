Attribute VB_Name = "Arrays2"
Public Function ReturnArray(arr, Optional ByRef Application_Caller As Object)
    If IsMissing(Application_Caller) Then
        CallerRows = UBound(arr)
        CallerCols = 1
    Else
        CallerRows = Application_Caller.Rows.Count
        CallerCols = Application_Caller.columns.Count
    End If

    ReDim Result(1 To CallerRows, 1 To CallerCols)
    For RowNdx = 1 To CallerRows
        For ColNdx = 1 To CallerCols
            n = n + 1
            Result(RowNdx, ColNdx) = GetItem(arr, n)
        Next ColNdx
     Next RowNdx
    
    ReturnArray = Result

End Function

Function ExcludeEmpty(ByRef arr As Variant)
    Dim Results As Variant
    ReDim Results(1 To 1)
    
    For x = LBound(arr) To UBound(arr)
        If arr(x) <> vbEmpty Then Results(UBound(Results)) = arr(x)
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

Function GetItem(arr, Index, Optional Default As String = "")
    On Error GoTo ErrHandler
    GetItem = arr(Index)
    Exit Function
    
ErrHandler:
    GetItem = Default

End Function
