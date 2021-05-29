### Function To Remove Character From End

```VBA
Private Function RemoveCharacterFromEnd(str As String, NumberOfCharacter As Integer) As String
    RemoveCharacterFromEnd = Left(str, Len(str) - NumberOfCharacter)
End Function
```

### Regex to Find All Matches in String
- Pattern – The pattern you are going to use for matching against the string.
- IgnoreCase – If True, then the matching ignores letter case.
- Global – If True, then all the matches of the pattern in the string are found. If False then only the first match is found.
- MultiLine – If True, pattern matching happens across line breaks.

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
