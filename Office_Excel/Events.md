### Run Macro Every time Worksheet Change
```vba
'Remove Case Sensitivity
  Option Compare Text

Private Sub Worksheet_Change(ByVal Target As Range)

'Determine if change was made to cell G7
  If Not Intersect(Target, Range("G7")) Is Nothing Then
    
    'Determine if the work "yes" is contained within cell G7
      If InStr(1, Range("G7"), "Yes") > 0 Then
        Range("G9").Font.Color = Range("F5").Font.Color
      Else
        Range("G9").Font.Color = Range("G9").Interior.Color
      End If
  
  End If

End Sub
```
