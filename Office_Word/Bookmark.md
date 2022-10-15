### Get Range Between Two BookMarks

```vba
Public Sub Test()
    Dim str As String
    str = "Vivek"
    Dim orng As Range
    Set orng = ActiveDocument.Range
    orng.Start = orng.Bookmarks("Start").Range.End + 1
    orng.End = orng.Bookmarks("End").Range.Start - 1
    orng.Text = str
    Debug.Print , orng.Text
    
    'the bookmarks are still present and the range can be reselected
    orng.Start = ActiveDocument.Bookmarks("Start").Range.End + 1
    orng.End = ActiveDocument.Bookmarks("End").Range.Start - 1
    orng.Select
End Sub
```
