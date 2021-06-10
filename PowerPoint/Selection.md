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
