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
Set activeSlide = Application.ActiveWindow.View.Slide
```
