Attribute VB_Name = "CollectionSort"
Function GetSortKeyByName(procName As String, V As Variant) As Variant
    GetSortKeyByName = CallByName(New SortKeys, procName, VbMethod, V)
End Function

Sub CSort(c As Collection, SortKeyFunction As String)
    Dim i As Long, j As Long
    For i = 1 To c.Count
        For j = c.Count To i Step -1
            If GetSortKeyByName(SortKeyFunction, c(i)) _
                > GetSortKeyByName(SortKeyFunction, c(j)) Then
                    CollectionSwap c, i, j
            End If
        Next j
    Next i
End Sub
Sub CollectionSwap(c As Collection, Index1 As Long, Index2 As Long)
    Dim Item1 As Variant, Item2 As Variant
    
    If IsObject(c.Item(Index1)) Then
        Set Item1 = c.Item(Index1)
    Else
        Let Item1 = c.Item(Index1)
    End If
    
    If IsObject(c.Item(Index2)) Then
        Set Item2 = c.Item(Index2)
    Else
        Let Item2 = c.Item(Index2)
    End If
    
    c.Add Item1, after:=Index2
    c.Remove Index2
    c.Add Item2, after:=Index1
    c.Remove Index1
End Sub
