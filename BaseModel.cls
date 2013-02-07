VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private 系able As String
Private 狄ields As Variant
Private 狄ilters As String

Public Property Get table() As String
    table = 系able
End Property
Public Property Let table(value As String)
    系able = value
End Property

Public Property Get filters() As String
    filters = 系able
End Property
Public Property Let filters(value As String)
    狄ilters = value
End Property

