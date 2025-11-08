### Update Word Table
```vba
Sub UpdateTable()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument
    
    Dim table As table
    Set table = wdDoc.Tables(1)
    Debug.Print , table.Rows.Count
    Debug.Print , table.Columns.Count
    Debug.Print , table.Cell(1, 1).range.Text
    
    'Update Specific Value
    table.Cell(2, 2).range.Text = "Calculator"
    
    'Add New Row
    table.Rows.Add
    table.Rows(table.Rows.Count).Cells(1).range.Text = table.Rows.Count - 1
    table.Rows(table.Rows.Count).Cells(2).range.Text = "Test"
    table.Rows(table.Rows.Count).Cells(3).range.Text = 10
End Sub
```

### Copy Table From Excel to Active Document
```vba
Sub CopyTableFromExcel()
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    
    'Get the running word application
    Set wdApp = GetObject(, "Word.Application")

    'select the open document you want to paste into
    Set wdDoc = wdApp.ActiveDocument
    
    Sheet1.ListObjects("Table1").Range.Copy
    
    
    'select the word range you want to paste into
    wdDoc.Bookmarks("Table1").Select

    'and paste the clipboard contents
    wdApp.Selection.Paste
End Sub
```
