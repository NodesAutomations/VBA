Attribute VB_Name = "Basic"
'@Folder("VBAProject")
Option Explicit

Sub API()

    Dim request As New WinHttpRequest
    request.Open "Get", "https://api.nationalize.io?name=michael"
    
    request.Send
    
    If request.Status <> 200 Then
        Exit Sub
    End If
    
    Dim response As Object
    Set response = JsonConverter.ParseJson(request.ResponseText)
    
    Debug.Print response("name")
    
    Dim countries As Collection
    Set countries = response("country")
    
    Dim country As Dictionary
    For Each country In countries
        Debug.Print country("country_id"), country("probability")
    Next
End Sub

