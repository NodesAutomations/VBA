### Copy Text To Windows ClipBoard
```vba
Function Clipboard$(Optional s$)
    'Cast to variant for 64-bit VBA support
    Dim v: v = s
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(s): .setData "text", v
            Case Else:   Clipboard = .GetData("text")
            End Select
        End With
    End With
End Function
```
