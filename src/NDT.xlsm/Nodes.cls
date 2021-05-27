VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Nodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private internalCollection As Collection

Private Sub Class_Initialize()
    Set internalCollection = New Collection
End Sub

Public Sub AddNode(n As Node, Optional key, Optional Before, Optional After)
    internalCollection.Add n, key, Before, After
End Sub

Public Function Item(Index) As Node
Attribute Item.VB_UserMemId = 0
    Item = internalCollection.Item(Index)
End Function

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
         Set NewEnum = internalCollection.[_NewEnum]
End Function
