### Function To Remove Character From End

```VBA
Private Function RemoveCharacterFromEnd(str As String, NumberOfCharacter As Integer) As String
    RemoveCharacterFromEnd = Left(str, Len(str) - NumberOfCharacter)
End Function
```
