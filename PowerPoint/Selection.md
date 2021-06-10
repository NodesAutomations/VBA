### Check if Any Object is Selected
```vba
Sub CheckSelectionType()
    Debug.Print , ActiveWindow.Selection.Type
    '0=Nothing Selected
    '1=Slide Selected
    '2=Shapes Selected
    '3=TextRange Selected
End Sub
``` 
### LoopThrough Selected Shapes
```vba
Public Sub PrintSelectedShapes()
    Dim shape As shape
    For Each shape In ActiveWindow.Selection.ShapeRange
        Debug.Print , shape.TextFrame.TextRange.Text
    Next
End Sub
```
