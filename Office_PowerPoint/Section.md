### Display Section Name
```vba
Sub DisplaySectionName()
    'Active Slide
    Dim activeSlide As Slide
    Set activeSlide = ActiveWindow.View.Slide
    MsgBox ActivePresentation.SectionProperties.Name(activeSlide.sectionIndex)
End Sub

```

### Loop Through Each Section
```vba
    Dim i As Integer
    For i = 1 To ActivePresentation.SectionProperties.Count
        Debug.Print ActivePresentation.SectionProperties.Name(i)
    Next
```
