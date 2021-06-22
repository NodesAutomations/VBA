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
### Clear Table
```vba
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.ActiveSheet.ListObjects("SaleTable")
    
    'Delete all table rows except first row
    With tbl.DataBodyRange
        If .Rows.Count > 1 Then
            .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
        End If
    End With
    tbl.DataBodyRange(1, 1) = ""
    tbl.DataBodyRange(1, 2) = ""
    tbl.DataBodyRange(1, 3) = ""
    tbl.DataBodyRange(1, 8) = 0
```
