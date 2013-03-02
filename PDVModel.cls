VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PDVModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class PDVModel created on '----------------------------------------------------------01/03/2013 11:43:46
Private ¨GE As String
Private ¨Custcode As String

' Base model definition ------------------------------------------------------------------------
Private ¨table As String
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

Private Sub Class_Initialize()

    ReDim flds(0 To 1, 0 To 2)
    Set Me.fields = flds

    Me.table = "VW_AUX_PDV_CUST_CODE_LAST"

    Me.filters = "" & _
         "WHERE" & _
         "    CANAL_N3 IN ('LOJAS PREMIUM', 'DEALER', 'SMALL RETAIL') AND " & _
         "    STATUS IN ('CREDENCIADO ATIVO', 'CREDENCIADO INATIVO', 'EM CREDENCIAMENTO')"

    flds(0, 0) = "Custcode"
    flds(0, 1) = "CUST_CODE"
    flds(0, 2) = "String"
    flds(1, 0) = "GE"
    flds(1, 1) = "GRUPO_ECONOMICO"
    flds(1, 2) = "String"

    Me.fields = flds

End Sub



