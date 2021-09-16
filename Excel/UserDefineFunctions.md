### Function to Extract Number From String
```vba
'Extract Number From string
Function GetNumeric(text As String) As Integer
    Dim stringLength As Integer
    Dim result As Integer
    stringLength = Len(text)
    Dim i As Integer
    For i = 1 To stringLength
        If IsNumeric(Mid(text, i, 1)) Then result = result & Mid(text, i, 1)
    Next i
    GetNumeric = result
End Function
```
```vba
Function CONCATENATEMULTIPLE(Ref As Range, Separator As String) As String
    Dim Cell As Range
    Dim Result As String
    For Each Cell In Ref
        Result = Result & Cell.Value & Separator
    Next Cell
    CONCATENATEMULTIPLE = Left(Result, Len(Result) - Len(Separator))
End Function
```
