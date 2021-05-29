### Function To Remove Character From End

```VBA
Private Function RemoveCharacterFromEnd(str As String, NumberOfCharacter As Integer) As String
    RemoveCharacterFromEnd = Left(str, Len(str) - NumberOfCharacter)
End Function
```

### Regex to Find All Matches in String

```vba
 Dim regx As New RegExp
    With regx
        .Pattern = " (?!\bFrom\s)(.+)(?=:)"
        .Global = True
        .Multiline = True
    End With
    
    Dim m As Match
    Dim c As MatchCollection

    Set c = regx.Execute(chatData)

    For Each m In c
        Debug.Print , m.Value
    Next
```
