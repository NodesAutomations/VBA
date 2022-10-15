### Find Text in Document
```vba
Sub FindText()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument

    Dim MyRange As Range
    Set MyRange = ActiveDocument.Range
    MyRange.Find.Execute FindText:="<<Test1>>"
 
    MsgBox "Position = " & MyRange.Text
    MyRange.Paste
   
End Sub
```
### Find text in Table  
```vba
Sub findTextInsideTable()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument
    Dim tbl As Table
    Dim MyRange As Range
    For Each tbl In wdDoc.Tables
        Set MyRange = tbl.Range
    
        MyRange.Find.Execute FindText:="<<Test6>>"
        
        If MyRange.Start > 0 Then
            MsgBox "Position = " & MyRange.Start
            'MyRange.Paste
        End If
    Next
    
End Sub
```

