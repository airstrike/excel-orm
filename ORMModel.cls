VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORMModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ¨table As String
Private ¨fields As Variant
Private ¨filters As String
Private ¨def As String
Dim flds As Variant

Private Sub class_initialize()
    ReDim flds(0 To 2, 0 To 0)
    Set Me.fields = flds
    

End Sub

Sub make(classname As String)
    Me.§ ParseBaseModel()
    Me.make_fields
    'For x = LBound(Me.fields) To UBound(Me.fields)
    '    'make sure this is correct
    '    Me.§ "Private " & Me.fields(x)(0) & " As " & Me.fields(x)(1).vb_type
    'Next

    Debug.Print Me.def
    
    'create fields again in the created module after making it, probably on __init__?
End Sub

Sub make_fields()
    Dim s As String
    Me.§ vbCrLf & "Private Sub class_initialize()" & vbCrLf
    For x = LBound(Me.fields) To UBound(Me.fields)
        Call Me.§("Private ¨" & Me.fields(x, 0) & " As " & Me.fields(x, 2), -1)
        
        For i = 0 To 2
        Me.§ "    Me.fields(" & x & ", " & i & ") = " & Me.fields(x, i)
        Next
        
    Next
    
    Me.§ vbCrLf & "End sub" & vbCrLf
    
End Sub

Public Property Set fields(flds As Variant)
    Call CopyArray(¨fields, flds)
End Property


Public Property Get fields() As Variant
    fields = ¨fields
End Property

Public Property Let fields(value As Variant)
    ¨fields = value
End Property

Public Property Get table() As String
    table = ¨table
    
End Property

Public Property Let table(value As String)
    ¨table = value

End Property

Public Property Get filters() As String
    filters = ¨filters
    
End Property
Public Property Let filters(value As String)
    ¨filters = value
    
End Property

Property Get field_(index As Long) As Variant
    field = anArray(index)
End Property
 
Property Let field_(index As Long, inValue As Variant)
    Me.fields(index) = inValue
End Property

Public Function add_field(name As String, ByVal field As Object) As Variant
    Dim fld As Variant, flds As Variant
    ReDim fld(0 To 2)
    Dim pos As Long

    fld(0) = name
    fld(1) = field.db_column
    fld(2) = field.vb_type
    
    If Not IsArrayAllocated(Me.fields) Then
        pos = 0
        ReDim flds(0 To 0, 0 To 2)
    Else
        pos = UBound(Me.fields) + 1
        flds = ExpandArray(Me.fields, 1, 1, 0)
    End If
    
    If Not IsArrayAllocated(flds) Then
        ReDim flds(0 To pos, 0 To 2)
    Else
        flds(pos, 0) = fld(0)
        flds(pos, 1) = fld(1)
        flds(pos, 2) = fld(2)

    End If
    Me.fields = flds
        
End Function

'------------------------------------------------'
' Shortcuts for adding definitions to this class '
'------------------------------------------------'
Sub §(Optional s As String = vbCrLf, Optional pos As Long = 1)
    If pos = 1 Then
        Me.def = Me.def & vbCrLf & s
    Else
        Me.def = s & vbCrLf & Me.def
    End If
End Sub

Public Property Get def() As String
    def = ¨def
    
End Property
Public Property Let def(value As String)
    ¨def = value
    
End Property
