### Check if worksheet Exist
```vba
Function IsWorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    IsWorksheetExists = Not sht Is Nothing
End Function
```

### Selection Change Event
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
   If Target.Cells.Count > 0 Then
        If Not Application.Intersect(Target, DataSheet.ListObjects("DataTable").Range) Is Nothing Then
            Call UI.ReloadRibbon(0)
        End If
    End If
End Sub

```

