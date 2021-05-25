VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Require refference for Microsoft Scripting Runtime Library
Public TaskTitle As String
Public Dependency As Scripting.Dictionary

Private Sub Class_Initialize()
    Set Dependency = New Scripting.Dictionary
End Sub

Public Sub AddDependency(n As Node)
    If Not Dependency.Exists(n.TaskTitle) Then
        Dependency.Add n.TaskTitle, n
    End If
End Sub

Public Property Get Self() As Object
    Set Self = Me
End Property

Public Property Get TaskNumber()
    TaskNumber = Val(TaskTitle)
End Property

Public Property Get UnnumberedTaskTitle()
    UnnumberedTaskTitle = Mid(TaskTitle, InStr(1, TaskTitle, ".") + 1)
End Property

Public Sub DumpStatus()
    Debug.Print "Å°" & Me.TaskTitle & " Depends on:"
    Dim k As Variant
    For Each k In Dependency.Keys
        Debug.Print Dependency.Item(k).TaskTitle
    Next
End Sub

Public Function GetDependency() As Collection
    Dim c As Collection
    Set c = New Collection
    Dim k As Variant
    For Each k In Dependency.Keys
        c.Add Dependency.Item(k)
    Next
    CollectionSort.CSort c, "SortKey1"
    Set GetDependency = c
End Function

Private Sub Class_Terminate()
    Dependency.RemoveAll
End Sub
