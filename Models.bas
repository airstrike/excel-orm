Attribute VB_Name = "Models"
Sub ModelPDV()
    Dim m As ORMModel
    Set m = New ORMModel
    
    m.table = "VW_AUX_PDV_CUST_CODE_LAST"
    m.filters = filters("SEGMENTO__iexact", "CONSUMER", _
                        "CANAL_N3__in", ¨("LOJAS PREMIUM", "DEALER", "SMALL RETAIL"), _
                        "STATUS__in", ¨("CREDENCIADO ATIVO", "CREDENCIADO INATIVO", "EM CREDENCIAMENTO") _
    )
    
    m.add_field name:="Custcode", field:=CharField("CUSTCODE")
    m.add_field name:="GE", field:=CharField("GRUPO_ECONOMICO")
    
    m.make "PDV"
    
End Sub
