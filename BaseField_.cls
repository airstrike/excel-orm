VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseField_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ¨vb_type As String
Public ¨db_column As String

Private Sub class_initialize()

End Sub

Public Property Get vb_type() As String
    vb_type = ¨vb_type
End Property
Public Property Let vb_type(value As String)
    ¨vb_type = value
End Property

Public Property Get db_column() As String
    db_column = ¨db_column
End Property
Public Property Let db_column(value As String)
    ¨db_column = value
End Property
