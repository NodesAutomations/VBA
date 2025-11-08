### Add New Table

```vba
Sub AddNewTable()
    Dim activeSlide As slide
    Set activeSlide = ActiveWindow.View.slide

    'Insert Table
    Dim tableShape As Shape
    Dim i As Integer, j As Integer
    Dim rows As Integer, columns As Integer
    rows = 3
    columns = 10
    Set tableShape = activeSlide.Shapes.AddTable(rows, columns)
    For i = 1 To rows
        For j = 1 To columns
            tableShape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text = i & j
        Next
    Next
End Sub
```
### Update Table
```vba
Sub UpdateTable()
    'Check If Some Shape is Selected
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox ("No Shape Selected")
        Exit Sub
    End If
    
    Dim shapeRange As shapeRange
    Set shapeRange = ActiveWindow.Selection.shapeRange
    
    'Check if Two Object Is Selected
    If shapeRange.Count <> 1 Then
        MsgBox ("Select Only One Shape")
        Exit Sub
    End If

    'Get Active Slide
    Dim activeSlide As slide
    Set activeSlide = ActiveWindow.View.slide
    
    'Store Shape Objects
    Dim shape As shape
    Set shape = shapeRange(1)
    
    If Not shape.Type = msoTable Then
        Exit Sub
    End If
    
    'Update Shape
    
    Dim i As Integer, j As Integer
    Dim rows As Integer, columns As Integer
    rows = shape.Table.rows.Count
    columns = shape.Table.columns.Count
   
    For i = 2 To rows
        For j = 1 To columns
            shape.Table.Cell(i, j).shape.TextFrame.TextRange.Text = "0"
        Next
    Next
    
    
End Sub
```

