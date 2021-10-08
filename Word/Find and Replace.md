### Code to Find Specific Item then Replace it
```vba
Private Sub FindAndReplace(wdDoc As Word.Document, findString As String, replaceString As String)

    'Update Main Content
    With wdDoc.Content.Find
        'Clear previous Formatting Settings
        .ClearFormatting
        .Replacement.ClearFormatting
        'What to Find
        .Text = findString
        'Only Find Text if Alignment Match
        '.ParagraphFormat.Alignment = wdAlignParagraphCenter
        'What to Replace
        .Replacement.Text = replaceString
        .Wrap = wdFindContinue
        'What to do on Match Find
        .Execute Replace:=wdReplaceAll
    End With
 
End Sub

```
