### Find Last USed Row Column
```vba
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.range("A" & ws.Rows.Count).End(xlUp).Row
```
