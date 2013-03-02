Attribute VB_Name = "TypeCreation"
Private Function JoinAttrs(attrs As Variant) As String
    JoinAttrs = ""
    For x = LBound(attrs) To UBound(attrs)
        JoinAttrs = JoinAttrs & "    " & attrs(0, x) & " as " & attrs(1, x) & vbCrLf
    Next
    
End Function

Sub TestCreateType()
    Dim attrs As Variant
    ReDim attrs(0 To 1, 0 To 2)
    attrs(0, 0) = "Name"
    attrs(1, 0) = "String"
    attrs(0, 1) = "Bar"
    attrs(1, 1) = "String"
    
    CreateType typename:="TestType", attrs:=attrs

End Sub

Sub CreateType(typename As String, Optional attrs As Variant)
    'If IsMissing(bases) Then
    '    ReDim bases(0 To 0)
    '    bases(0) = ""
    'End If
    
    If IsMissing(attrs) Then
        ReDim attrs(0 To 1, 0 To 0)
        attrs(0, 0) = ""
        attrs(0, 1) = ""
    End If
    
    CreateType_ typename, attrs ', bases, attrs
End Sub

Private Sub CreateType_(typename As String, attrs As Variant)

    'GoTo SkipDeletion
    Dim startLine, endLine, _
        i, idxStart, idxEnd, found

    With Application.VBE.ActiveVBProject.VBComponents("Types").CodeModule
        If .Find("Public Type " & typename, 1, 1, .CountOfDeclarationLines, 99999, True) Then
            found = False: idxStart = 0: idxEnd = 0
            startLine = 0: endLine = 0

            'Find the existing module definition and remove it
            For i = 1 To .CountOfDeclarationLines
                idxStart = InStr(1, .Lines(i, 1), "Public Type " & typename)
                If idxStart > 0 Then
                    found = True
                    startLine = i
                End If

                If found = True Then
                    idxEnd = InStr(1, .Lines(i, 1), "End Type")
                    If idxEnd > 0 Then
                        endLine = i
                    End If
                End If
            Next
        End If
        If startLine > 0 And endLine > 0 Then
            .DeleteLines startLine, endLine - startLine + 1
        End If
    End With

SkipDeletion:
    'def is the type definition
    Dim def As String
    def = ""
    def = "Public Type " & typename & vbCrLf
    
    If Not IsArrayEmpty(attrs) Then 'LBound(attrs) = 0 And Not UBound(attrs) = 0 Then
        def = def & JoinAttrs(attrs) & vbCrLf
    End If
    
    def = def & "End Type" & vbCrLf
    
    'Debug.Print def
    
    'Exit Sub
    Application.VBE.ActiveVBProject.VBComponents("Types").CodeModule.AddFromString (def)

End Sub
