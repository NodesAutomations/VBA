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
