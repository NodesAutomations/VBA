### Display Section Name
```vba
Sub DisplaySectionName()
    'Active Slide
    Dim activeSlide As Slide
    Set activeSlide = ActiveWindow.View.Slide
    MsgBox ActivePresentation.SectionProperties.Name(activeSlide.sectionIndex)
End Sub

```
