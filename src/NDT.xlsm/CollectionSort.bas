Attribute VB_Name = "CollectionSort"
Function GetSortKeyByName(procName As String, V As Variant) As Variant
    GetSortKeyByName = CallByName(New SortKeys, procName, VbMethod, V)
End Function

Sub CSort(C As Collection, SortKeyFunction As String)
    Dim i As Long, j As Long
    For i = 1 To C.Count
        For j = C.Count To i Step -1
            If GetSortKeyByName(SortKeyFunction, C(i)) _
                > GetSortKeyByName(SortKeyFunction, C(j)) Then
                    CollectionSwap C, i, j
            End If
        Next j
    Next i
End Sub
Sub CollectionSwap(C As Collection, Index1 As Long, Index2 As Long)
    Dim Item1 As Variant, Item2 As Variant
    
    If IsObject(C.Item(Index1)) Then
        Set Item1 = C.Item(Index1)
    Else
        Let Item1 = C.Item(Index1)
    End If
    
    If IsObject(C.Item(Index2)) Then
        Set Item2 = C.Item(Index2)
    Else
        Let Item2 = C.Item(Index2)
    End If
    
    C.Add Item1, after:=Index2
    C.Remove Index2
    C.Add Item2, after:=Index1
    C.Remove Index1
End Sub
