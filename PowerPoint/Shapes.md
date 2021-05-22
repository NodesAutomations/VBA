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

### Copy All Object From One Slide to Another Slide

```VBA
Sub CopyAllObjectFromFirstSlideToSecond()
    Dim curSlide As Slide
    Dim curShape As Shape
 
    With ActivePresentation
    
        'Copy All Object
        .Slides(1).Shapes.Range.Copy
    
        ''Copy 1st and 3rd Object
        '.Slides(1).Shapes.Range(Array(1, 3)).Copy
    
        ''Only Copy 3rd object
        '.Slides(1).Shapes.Range(Array(3)).Copy
    
        .Slides(2).Shapes.Paste

    End With
End Sub
```

### Change Name Of Shape
```VBA
Sub ChangeId()
    ActivePresentation.Slides(1).Shapes(1).Name = "test"
    Debug.Print , ActivePresentation.Slides(1).Shapes(1).Name
End Sub

```



