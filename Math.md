### Random Index Generation

```vba
Private Function GiveRandomIndex(maxIndex As Integer) As Integer
    GiveRandomIndex = Int(Rnd() * maxIndex) + 1
End Function
```
