### List All Shapes

```VBA
Sub ListAllObjects()
    Dim curSlide As Slide
    Dim curShape As Shape
    For Each curSlide In ActivePresentation.Slides
        For Each curShape In curSlide.Shapes
            Debug.Print , curShape.Name
        Next curShape
    Next curSlide
End Sub
```