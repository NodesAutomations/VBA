### Function to Check if String Contain SubString
- Instring function basically return postion of substring in string
- If substring don't exist it returns 0
```vba
If InStr(string, substring) > 0 then
debug.print "String Contain substring"
else
debug.Pring "String don't contain any substring"
end if
```


### Function Split String Into Array by deliminator
```vba
   Dim StringValue As String
    StringValue ="Test1,Test2"
    StringValue = Trim(StringValue)
    
    Dim Data()  As String
    Data = Split(StringValue, ",")
```
### Function To Remove Character From End

```VBA
Private Function RemoveCharacterFromEnd(str As String, NumberOfCharacter As Integer) As String
    RemoveCharacterFromEnd = Left(str, Len(str) - NumberOfCharacter)
End Function
```

### Regex to Find All Matches in String

#### Properties
- Pattern – The pattern you are going to use for matching against the string.
- IgnoreCase – If True, then the matching ignores letter case.
- Global – If True, then all the matches of the pattern in the string are found. If False then only the first match is found.
- MultiLine – If True, pattern matching happens across line breaks.

#### Methods
- Test – Searches for a pattern in a string and returns True if a match is found.
- Replace – Replaces the occurrences of the pattern with the replacement string.
- Execute – Returns matches of the pattern against the string.

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
