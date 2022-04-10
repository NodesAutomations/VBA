### Get Active SlideIndex
```vba
Dim activeSlideIndex as Integer
'For Normal Mode
activeSlideIndex = ActiveWindow.View.Slide.SlideIndex
'For Presentation Mode
activeSlideIndex = ActivePresentation.SlideShowWindow.View.Slide.SlideIndex
```
### Get Active Slide
```vba
Dim activeSlide As Slide
'For Normal Mode
Set activeSlide = ActiveWindow.View.Slide
'For Presentation Mode
Set activeSlide = ActivePresentation.SlideShowWindow.View.Slide
```

### Loop Through All Slides
```vba
    Dim slide As slide
    For Each slide In ActivePresentation.Slides
       Debug.Print slide.SlideIndex
    Next
```
### Get Slide Title
```vba
Sub Test()
    Dim slide As slide
 
    For Each slide In ActivePresentation.Slides
        If slide.Shapes.HasTitle Then
            Debug.Print slide.SlideNumber, slide.Shapes.Title.TextFrame.TextRange.Text
        End If
    Next
End Sub
```
