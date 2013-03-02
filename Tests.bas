Attribute VB_Name = "Tests"
'Call TestHash

'---------------------------------------------
' Begin tests
'---------------------------------------------
Sub TestHash()
    Dim ht As HashTable
    Set ht = New HashTable
    
    ht.SetSize 100 'any number > 0 is valid here
    ht("Foo") = "Bar"
    If ht("Foo") = "Bar" Then Exit Sub

End Sub
