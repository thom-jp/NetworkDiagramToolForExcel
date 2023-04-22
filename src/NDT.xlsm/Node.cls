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
Public ShapeObjectName As String
Public TaskListRange As Range
Public Dependency As Scripting.Dictionary

Private Sub Class_Initialize()
    Set Dependency = New Scripting.Dictionary
End Sub

Public Function FindShape() As Shape
    On Error GoTo Err_Handler
    If ShapeObjectName <> "" Then
        Set FindShape = DrawSheet.Shapes(ShapeObjectName)
    End If
    On Error GoTo 0
Exit Function
Err_Handler:
    Set FindShape = Nothing
End Function

Public Sub AddDependency(n As Node)
    If Not Dependency.Exists(n.TaskTitle) Then
        Dependency.Add n.TaskTitle, n
    End If
End Sub

Public Property Get Self() As Object
    Set Self = Me
End Property

Public Property Get TaskNumber()
    TaskNumber = val(TaskTitle)
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

Public Function GetDependency() As Nodes
    Dim c As Nodes
    Set c = New Nodes
    Dim k As Variant
    For Each k In Dependency.Keys
        c.AddNode Dependency.Item(k)
    Next
    c.Sort
    Set GetDependency = c
End Function

Private Sub Class_Terminate()
    Dependency.RemoveAll
End Sub
