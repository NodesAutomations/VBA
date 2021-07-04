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
    wdAp
p.Selection.Paste
End Sub
```
