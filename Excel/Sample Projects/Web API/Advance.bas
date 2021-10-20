Attribute VB_Name = "Advance"
'@Folder("VBAProject")
Option Explicit

Sub API()

    Dim request As New WinHttpRequest
    request.Open "Get", "https://api.nationalize.io?name=michael"
    
    request.Send
    
    If request.Status <> 200 Then
        Exit Sub
    End If
    Debug.Print request.ResponseText
End Sub


