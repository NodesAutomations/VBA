### Shape Formatting
```vba
Sub Create_A_Shape()

'PURPOSE:Create a Text Box Shape and Reformat it
'SOURCE: www.TheSpreadsheetGuru.com

Dim Sld As Slide
Dim Shp As Shape

'ERROR HANDLING
    If ActivePresentation.Slides.Count = 0 Then
        MsgBox "You do not have any slides in your PowerPoint project."
        Exit Sub
    End If

Set Sld = Application.ActiveWindow.View.Slide

'Create shape with Specified Dimensions and Slide Position
    Set Shp = Sld.Shapes.AddShape(Type:=msoShapeRectangle, _
        Left:=24, Top:=65.6, Width:=672, Height:=26.6)

'FORMAT SHAPE
    'Shape Name
        Shp.Name = "My Header"
    
    'No Shape Border
        Shp.Line.Visible = msoFalse
    
    'Shape Fill Color
        Shp.Fill.ForeColor.RGB = RGB(184, 59, 29)
    
    'Shape Text Color
        Shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    
    'Text inside Shape
        Shp.TextFrame.TextRange.Characters.Text = "[Header]"
    
    'Center Align Text
        Shp.TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    
    'Vertically Align Text to Middle
        Shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    
    'Adjust Font Size
        Shp.TextFrame2.TextRange.Font.Size = 14
    
    'Adjust Font Style
        Shp.TextFrame2.TextRange.Font.Name = "Arial"

End Sub
```
