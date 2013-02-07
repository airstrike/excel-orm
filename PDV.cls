VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub ModelPDV()
    m = New ORMModel
    
    m.table = "VW_AUX_PDV_CUST_CODE_LAST"
    m.filters = ORMFilter("SEGMENTO__iexact", "CONSUMER", _
                              "CANAL_N3__in", ¨("LOJAS PREMIUM", "DEALER", "SMALL RETAIL"), _
                              "STATUS__in", ¨("CREDENCIADO ATIVO", "CREDENCIADO INATIVO", "EM CREDENCIAMENTO") _
    )
    
    m.AddField name:="Custcode", field:=ORMCharField("CUSTCODE")
    m.AddField name:="GE", field:=ORMCharField("GRUPO_ECONOMICO")
    
    m.Register name:="PDV"
    
End Sub


Private ¨table As String
Private ¨filters As String
Private ¨Custcode As Object
Private ¨GE As Object

Public Property Get table() As String
    table = ¨table
End Property
Public Property Let table(value As String)
    ¨table = value
End Property

