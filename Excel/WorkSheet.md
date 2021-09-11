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

