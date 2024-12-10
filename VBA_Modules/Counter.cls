'@Folder("VBAProject")
Option Explicit

Private pStartingId As Integer
Private pIncrement As Integer
Private pId As Integer

Public Property Get Id() As Integer
    Id = pId
End Property

Public Property Get StartingId() As Integer
    StartingId = pStartingId
End Property

Public Property Get Increment() As Integer
    Increment = pIncrement
End Property

' Constructor with parameters
Public Sub SetProperty(Optional ByVal StartingId As Integer = 1, Optional ByVal Increment As Integer = 1)
    pStartingId = StartingId
    pIncrement = Increment
    Reset
End Sub

' Reset the counter to the starting ID
Public Sub Reset()
    pId = pStartingId
End Sub

' Return the current ID and increment by 1
Public Function CreateId() As Integer
    CreateId = CreateIdWithIncrement(1)
End Function

' Return the current ID and increment by a specified number
Public Function CreateIdWithIncrement(ByVal numberOfIncrement As Integer) As Integer
    Dim currentId As Integer
    currentId = pId
    pId = pId + pIncrement * numberOfIncrement
    CreateIdWithIncrement = currentId
End Function

' Override the ToString method
Public Function ToString() As String
    ToString = "Id:" & pId & ",Start:" & pStartingId & ",Inc:" & pIncrement
End Function

