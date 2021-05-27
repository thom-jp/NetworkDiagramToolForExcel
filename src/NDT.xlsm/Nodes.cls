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

Public Sub AddNode(n As Node, Optional key As String, Optional Before, Optional After)
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

Public Function Exists(ByVal key As String) As Boolean
    On Error GoTo Error_Handler:
        Call internalCollection.Item(key)
    On Error GoTo 0
    Exists = True
Exit Function
Error_Handler:
    If Err.Number = 5 Then
        Exists = False
    Else
        Err.Raise Err.Number
    End If
End Function

Public Sub Sort()
    Dim i As Long, j As Long
    For i = 1 To internalCollection.Count
        For j = internalCollection.Count To i Step -1
            Dim n As Node
            If Val(internalCollection(i).TaskTitle) > Val(internalCollection(j).TaskTitle) Then
                    SwapItems i, j
            End If
        Next j
    Next i
End Sub

Private Sub SwapItems(Index1 As Long, Index2 As Long)
    Dim Item1 As Variant, Item2 As Variant
    
    If IsObject(internalCollection.Item(Index1)) Then
        Set Item1 = internalCollection.Item(Index1)
    Else
        Let Item1 = internalCollection.Item(Index1)
    End If
    
    If IsObject(internalCollection.Item(Index2)) Then
        Set Item2 = internalCollection.Item(Index2)
    Else
        Let Item2 = internalCollection.Item(Index2)
    End If
    
    internalCollection.Add Item1, After:=Index2
    internalCollection.Remove Index2
    internalCollection.Add Item2, After:=Index1
    internalCollection.Remove Index1
End Sub

