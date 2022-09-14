### Copy Cell with Images to Same or antoher Sheet
```vba
Sub Test()
    Dim sourceRange As Range
    Set sourceRange = Sheet1.Range("E2:E12")
    
    Dim targetRange As Range
    Set targetRange = Sheet1.Range("E21")
    
    Dim i As Integer, rowId As Integer
    rowId = 1
    For i = 1 To sourceRange.Cells.Count
    
        'Code to Copy Cell with image
        Sheet1.Activate
        sourceRange.Cells(i, 1).Select
        Selection.Copy
        
        'Code to Paste Cell with image
        Sheet1.Activate
        targetRange.Cells(rowId, 1).Select
        Sheet1.Paste
        
        rowId = rowId + 1
    Next
    Application.CutCopyMode = False
End Sub
```

