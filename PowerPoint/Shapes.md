### Get Shape Using Name
```vba
ActivePresentation.Slides(1).Shapes("Player1Name").TextEffect.Text
```

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
### Remove Specific type of shapes from Slide
```VBA
Private Sub RemoveOldFolderObjects()
    Dim shape As shape
    Dim TotalShapes As Integer, i As Integer
    TotalShapes = ActivePresentation.Slides(1).Shapes.Count

    For i = TotalShapes To 1 Step -1
        Set shape = ActivePresentation.Slides(1).Shapes(i)
        If shape.Type = msoLine Then
            shape.Delete
        End If
    Next
 
End Sub
```
### Remove Shape From Active Slide

```VBA
Private Sub RemovePicture(PPSlide As PowerPoint.Slide, PictureName As String)
    Dim curShape As shape
    For Each curShape In PPSlide.Shapes
        If curShape.Name = PictureName Then
            curShape.Delete
        End If
    Next curShape
End Sub
```
### Function to Check If ShapeExist
```vba
Function IsShapeExists(PPSlide As PowerPoint.Slide, shapeName As String) As Boolean
    Dim shape As PowerPoint.shape
    For Each shape In PPSlide.Shapes
        If shape.Name = shapeName Then IsShapeExists = True
    Next shape
End Function
```
### Get Specific Shape From Active Slide

```VBA
Private Function GetShape(PPSlide As PowerPoint.Slide, PictureName As String) As shape
    Dim curShape As shape
    For Each curShape In PPSlide.Shapes
        If curShape.Name = PictureName Then
            Set GetShape = curShape
            Exit Function
        End If
    Next curShape
End Function
```

### Add Shapes
```vba
Private Function AddTextbox(PPSlide As PowerPoint.Slide, leftPos As Integer, topPos As Integer, width As Double, height As Double, data As String) As PowerPoint.shape
    'Add TextBox With Number
    Dim PPShape As PowerPoint.shape
    'Change AutoShapeType as per requirement
    Set PPShape = PPSlide.Shapes.AddShape(msoShapeRectangle, left:=leftPos, top:=topPos, width:=width, height:=height)
    Set AddTextbox = PPShape
End Function
```
### Add New Picture

```VBA
'Create New Pic On Same Position as Old One
    Dim PPShape As shape
    Set PPShape = PPSlide.Shapes.AddPicture(FileName:=PictureFilePath, _
                                            LinkToFile:=msoFalse, _
                                            SaveWithDocument:=msoTrue, _
                                            Left:=Picture.Left, _
                                            Top:=Picture.Top, _
                                            Width:=Picture.Width, _
                                            Height:=Picture.Height)

    PPShape.Name = Picture.Name
```
