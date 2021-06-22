### Code to Select Multiple Files and Replace Current List
```
Private Sub SelectFilesButton_Click()
    Dim sFilter As String
    sFilter = "PPT Files (*.pptx; *.pptm),*.pptx;*.pptm"
    
    ChDir ThisWorkbook.path
    
    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename(FileFilter:=sFilter, Title:="Please select an PowerPoint file", MultiSelect:=True)
    
    'Code to Exit Sub if no file selected
    If VarType(fileToOpen) < vbArray Then
        Exit Sub
    End If
    
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("FilePathTable")
    
    'Delete all table rows except first row
    With tbl.DataBodyRange
        If .Rows.Count > 1 Then
            .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
        End If
    End With
    
    'Clear out data from first table row (retaining formulas)
    tbl.DataBodyRange.Range("A1").Offset(0, 0).Value = ""
    
    Dim i As Integer
    Dim filePath As Variant
    
    For Each filePath In fileToOpen
        tbl.DataBodyRange.Range("A1").Offset(i, 0).Value = filePath
        i = i + 1
    Next filePath

    Set tbl = Nothing
End Sub
```

### Code to Select Multiple Files And Append to Current List
```vba
Private Sub SelectPPTs_Click()
 
    Dim sFilter As String
    sFilter = "PPT Files (*.pptx; *.pptm),*.pptx;*.pptm"
    
    ChDir ThisWorkbook.Path
    
    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename(FileFilter:=sFilter, Title:="Please select an PowerPoint file", MultiSelect:=True)
    
    'Code to Exit Sub if no file selected
    If VarType(fileToOpen) < vbArray Then
        Exit Sub
    End If
    
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.ActiveSheet.ListObjects("AudioReplacerTable")
    Dim columnId As Integer
    Dim rowCount As Integer
    
    rowCount = tbl.DataBodyRange.Rows.Count
    
    If tbl.DataBodyRange.Range("A1").Offset(0, columnId).Value = "" Then
        rowCount = rowCount - 1
    End If
    
    Dim i As Integer
    Dim filePath As Variant
    
    For Each filePath In fileToOpen
        tbl.DataBodyRange.Range("A1").Offset(rowCount + i, columnId).Value = filePath
        i = i + 1
    Next filePath

    Set tbl = Nothing
End Sub


```
