# Add Text
```vba
Public Function addMtext(ByVal originX As Double, ByVal originY As Double, ByVal width As Double, ByVal height As Double, ByVal text As String) As Object
    Dim corner(0 To 2) As Double
    'top left corner of text
    corner(0) = originX: corner(1) = originY: corner(2) = 0

    Set addMtext = cadModel.addMtext(corner, width, text)
    addMtext.height = height

End Function

```
