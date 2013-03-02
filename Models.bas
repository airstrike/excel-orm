Attribute VB_Name = "Models"
Sub MakePDV()
    Dim m As ORMModel
    Set m = New ORMModel
    
    m.table = "VW_AUX_PDV_CUST_CODE_LAST"
    '"SEGMENTO__iexact", "CONSUMER",
    m.filters = filter( _
                        "CANAL_N3__in", ¨("LOJAS PREMIUM", "DEALER", "SMALL RETAIL"), _
                        "STATUS__in", ¨("CREDENCIADO ATIVO", "CREDENCIADO INATIVO", "EM CREDENCIAMENTO") _
    )
    
    m.add_field Name:="Custcode", field:=CharField("CUST_CODE")
    m.add_field Name:="GE", field:=CharField("GRUPO_ECONOMICO")
    
    m.make "PDVModel"
    
End Sub

Sub PDVManager()
    Dim PDVCollection As Collection
    Dim MetaPDV As PDVModel
    Set MetaPDV = New PDVModel
    Dim columns As Variant
    ReDim columns(0 To UBound(MetaPDV.fields))
    
    strSql = ""
    Sql ("SELECT")
    For x = 0 To UBound(MetaPDV.fields)
        columns(x) = MetaPDV.fields(x, 1)
    Next x
    Sql (Join(columns, ", "))
    
    Sql ("FROM " & MetaPDV.table)
    Sql (MetaPDV.filters)

    z = 1
    Execute MapToModel:=MetaPDV
    

End Sub
