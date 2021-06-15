### Add Imgage From URL To Active Slide
```vba
Sub AddImageUsingURL()
    Dim imageURL As String
    imageURL = "https://pbs.twimg.com/media/E3D1W_wXIAIRJwO.png"
    
    'Get Active Slide
    Dim activeSlide As Slide
    Set activeSlide = ActiveWindow.View.Slide
    
    Dim PPShape As Shape
    Set PPShape = activeSlide.Shapes.AddPicture(imageURL, False, True, 0, 0, -1, -1)
                                            
End Sub
```
