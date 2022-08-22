### Syntax
```VBA
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.ActiveSheet.ListObjects("Table13")
    'First row , First column
    'tbl.Range.Cells(1, 1)
    'First row , First column above 1 row
    'tbl.Range.Cells(1, 1).Offset(-1,0)
```

### Get Active table

```VBA
Private Function GetActiveTable()
    On Error GoTo ERRORHANDLER
    Set GetActiveTable = ActiveCell.ListObject
    
ERRORHANDLER:
    If GetActiveTable Is Nothing Then
        ErrorMessage = "No Table Selected"
    End If

End Function
```
### Get Table Using only Name
```vba
Private Function GetTableObject(tableName As String) As Excel.ListObject
    On Error Resume Next
    Set GetTableObject = Application.range(tableName).ListObject
    On Error GoTo 0
    If GetTableObject Is Nothing Then
        Call Err.Raise(1004, ThisWorkbook.Name, "Table '" & tableName & "' not found!")
    End If
End Function

```
```vba
Public Function GetListObject(ByVal ListObjectName As String, Optional ParentWorksheet As Worksheet = Nothing) As Excel.ListObject
    On Error Resume Next

    If (Not ParentWorksheet Is Nothing) Then
        Set GetListObject = ParentWorksheet.ListObjects(ListObjectName)
    Else
        Set GetListObject = Application.range(ListObjectName).ListObject
    End If

    On Error GoTo 0                              'Or your error handler

    If (Not GetListObject Is Nothing) Then
        'Success
    ElseIf (Not ParentWorksheet Is Nothing) Then
        Call Err.Raise(1004, ThisWorkbook.Name, "ListObject '" & ListObjectName & "' not found on sheet '" & ParentWorksheet.Name & "'!")
    Else
        Call Err.Raise(1004, ThisWorkbook.Name, "ListObject '" & ListObjectName & "' not found!")
    End If

End Function

```
### Clear Table
```vba
 Private Sub ClearTableData(tbl As ListObject)
    'Delete all table rows except first row
    With tbl.DataBodyRange
        If .Rows.Count > 1 Then
            .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
        End If
    End With
    
    'Clear First Row
    tbl.range.Rows(2).Clear
End Sub
```
### Clear Table but keep formula's
```vba
Private Sub ClearTableData(tbl As ListObject)
    'Delete all table rows except first row
    If Not tbl.Range.Cells(2, 1).HasFormula Then
        If tbl.Range.Cells(2, 1) = "" Then
            tbl.Range.Cells(2, 1) = 1
        End If
    End If
    With tbl.DataBodyRange
        If Not tbl.DataBodyRange Is Nothing Then
            If .Rows.Count > 1 Then
                .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
            End If
        End If
    End With
    
    'Clear First Row but keep formula's
    Dim j As Integer
    For j = 1 To tbl.Range.Rows(2).Columns.Count
        If Not tbl.Range.Cells(2, j).HasFormula Then
            tbl.Range.Cells(2, j).Clear
        End If
    Next
 
End Sub
```

### Loop Through Table
```vba
 Dim tbl As ListObject
    Set tbl = AudioListSheet.ListObjects("AudioCategoryTable")
    
    Dim i As Integer
    For i = 1 To tbl.DataBodyRange.Rows.Count
        CategoryListBox.AddItem tbl.DataBodyRange.Cells(i, 1)
    Next
    
```
### Sort Table
```vba
Sub Sort()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects("myTable")
    Set rng = Range("myTable[Numbers]")
    
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
End Sub
```
Sort Multiple Column
```vba
Sub Sort()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects("myTable")
    Set rng1 = Range("myTable[First Name]")
    Set rng2 = Range("myTable[Last Name]")
    
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rng1, Order:=xlAscending
        .SortFields.Add Key:=rng2, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
End Sub
```
