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
