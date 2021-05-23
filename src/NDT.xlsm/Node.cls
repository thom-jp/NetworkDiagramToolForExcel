VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Sub DumpStatus()
    Debug.Print "��" & Me.TaskTitle & " Depends on:"
    For Each k In Dependency.Keys
        Debug.Print Dependency.Item(k).TaskTitle
    Next
End Sub

Private Sub Class_Terminate()
    Dependency.RemoveAll
End Sub