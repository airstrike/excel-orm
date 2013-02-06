VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORMModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Type ModelLibrary
    Dim Iter As Collection
    Set Iter = New Collection
    
    Sub Register(ModelName)
        If PresentInArray(Me.Models, "ModelName") Then Exit Sub
    End Sub

End Type

Public Type Model
    Name As String
    Definition As String
    Bases As Variant
    Namespace As String
    
End Type
