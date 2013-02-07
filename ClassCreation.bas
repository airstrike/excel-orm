Attribute VB_Name = "ClassCreation"
Private Function JoinAttrs(attrs As Variant) As String
    JoinAttrs = ""
    For x = LBound(attrs) To UBound(attrs)
        JoinAttrs = JoinAttrs & "Private " & attrs(0, x) & " as " & attrs(1, x) & vbCrLf
    Next
    
End Function

Sub TestCreateClass()
    Dim attrs As Variant
    ReDim attrs(0 To 1, 0 To 2)
    attrs(0, 0) = "Name"
    attrs(1, 0) = "String"
    attrs(0, 1) = "Bar"
    attrs(1, 1) = "String"
    
    CreateClass classname:="TestClass", attrs:=attrs

End Sub

Sub CreateClass(classname As String, Optional bases As Variant, Optional attrs As Variant)
    If IsMissing(bases) Then
        ReDim bases(0 To 0)
        bases(0) = ""
    End If
    
    If IsMissing(attrs) Then
        ReDim attrs(0 To 1, 0 To 0)
        attrs(0, 0) = ""
        attrs(0, 1) = ""
    End If
    
    CreateClass_ classname, bases, attrs
End Sub
Private Sub CreateClass_(classname As String, bases As Variant, attrs As Variant)
    Dim Class As VBComponent
    With Application.VBE.ActiveVBProject
        
        On Error GoTo CreateClassModule
        Set Class = .VBComponents(classname)
        On Error GoTo 0
    
        Call Class.CodeModule.DeleteLines(1, Class.CodeModule.CountOfLines)
        GoTo WriteClass
    
CreateClassModule:
    Set Class = .VBComponents.Add(vbext_ct_ClassModule)
    Class.name = classname
    
WriteClass:
        'def is the class definition
        Dim def As String
        def = ""
        If Not IsArrayEmpty(bases) Then 'LBound(bases) = 0 And Not UBound(bases) = 0 Then
            For x = LBound(bases) To UBound(bases)
                If bases(x) <> "" Then def = def & "Implements " & bases(x) & vbCrLf
            Next
        End If
        
        If Not IsArrayEmpty(attrs) Then 'LBound(attrs) = 0 And Not UBound(attrs) = 0 Then
            def = def & JoinAttrs(attrs) & vbCrLf
        End If
        
        def = def & vbCrLf
    
        Class.CodeModule.AddFromString (def)
    
    End With 'Application.VBE.ActiveVBProject

End Sub




