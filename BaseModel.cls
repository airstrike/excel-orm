VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private �table As String
Private �fields As Variant
Private �filters As String

Public Property Get table() As String
    table = �table
End Property
Public Property Let table(value As String)
    �table = value
End Property

Public Property Get filters() As String
    filters = �table
End Property
Public Property Let filters(value As String)
    �filters = value
End Property
