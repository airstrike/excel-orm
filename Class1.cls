VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Base model definition ------------------------------------------------------------------------
Private ¨hash As String
Private ¨timeout As String
Private ¨fields As Variant
Private ¨filters As String
Dim flds As Variant

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

' End of base model ------------------------------------------------------------------------

