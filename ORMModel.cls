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

Private Sub Class_Initialize()
    ReDim flds(0 To 2, 0 To 0)
    Set Me.fields = flds
    
End Sub

Sub make(classname As String)
    Me.§ ParseBaseModel()
    Me.make_initialization

    'For x = LBound(Me.fields) To UBound(Me.fields)
    '    'make sure this is correct
    '    Me.§ "Private " & Me.fields(x)(0) & " As " & Me.fields(x)(1).vb_type
    'Next

    'Debug.Print Me.def
    Call CreateClassFromString(classname, Me.def)
    
    'create fields again in the created module after making it, probably on __init__?
End Sub

Sub make_initialization()
    Me.§ vbCrLf & "Private Sub class_initialize()" & vbCrLf

    Me.§ "    ReDim flds(0 To " & UBound(Me.fields) & ", 0 To 2)"
    Me.§ "    Set Me.fields = flds"
    Me.§ ""
    Me.make_table
    Me.make_filters
    Me.make_fields
    
    Me.§ vbCrLf & "End sub" & vbCrLf
End Sub

Sub make_table()
    Me.§ "    Me.table = " & """" & Me.table & """" & vbCrLf
End Sub

Sub make_filters()
    Me.§ "    Me.filters = " & """" & Me.filters_to_string() & """" & vbCrLf
    'Me.§
    
End Sub

Function filters_to_string() As String
    Dim arr As Variant
    arr = Split(Me.filters, vbCrLf)
    filters_to_string = Join(arr, """" & " & _" & vbCrLf & "         " & """")

End Function

Sub make_fields()
    Dim s As String
    For x = LBound(Me.fields) To UBound(Me.fields)
        Call Me.§("Private ¨" & Me.fields(x, 0) & " As " & Me.fields(x, 2), -1)
        
        For i = 0 To 2
        Me.§ "    flds(" & x & ", " & i & ") = " & """" & Me.fields(x, i) & """"
        Next
        
    Next
    
    Me.§ vbCrLf & "    Me.fields = flds"
    
End Sub

Public Property Set fields(flds As Variant)
    Call CopyArray(¨fields, flds)
End Property


Public Property Get fields() As Variant
    fields = ¨fields
End Property

Public Property Let fields(Value As Variant)
    ¨fields = Value
End Property

Public Property Get table() As String
    table = ¨table
    
End Property

Public Property Let table(Value As String)
    ¨table = Value

End Property

Public Property Get filters() As String
    filters = ¨filters
    
End Property
Public Property Let filters(Value As String)
    ¨filters = Value
    
End Property

Property Get field_(Index As Long) As Variant
    field = anArray(Index)
End Property
 
Property Let field_(Index As Long, inValue As Variant)
    Me.fields(Index) = inValue
End Property

Public Function add_field(Name As String, ByVal field As Object) As Variant
    Dim fld As Variant, flds As Variant
    ReDim fld(0 To 2)
    Dim pos As Long

    fld(0) = Name
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
Public Property Let def(Value As String)
    ¨def = Value
    
End Property
