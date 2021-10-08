### Code to Find Specific Item then Replace it
```vba
Private Sub FindAndReplace(wdDoc As Word.Document, findString As String, replaceString As String)

    'Update Main Content
    With wdDoc.Content.Find
        .Text = findString
        .Replacement.Text = replaceString
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    'Update Header & Footer
    Dim section As Word.section
    Dim headerFooter As Word.headerFooter
    For Each section In wdDoc.Sections
        For Each headerFooter In section.Headers
            With headerFooter.Range.Find
                .Text = findString
                .Replacement.Text = replaceString
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Next
        
        For Each headerFooter In section.Footers
            With headerFooter.Range.Find
                .Text = findString
                .Replacement.Text = replaceString
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Next
    Next
End Sub

```
