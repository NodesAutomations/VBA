### Insert Image from URL
- Refer : https://excelmacroclass.blogspot.com/2022/01/insert-picture-from-url.html
Excel
```vba
  Sub InsertPictureFromURL()
 
      Dim url As String, x As Integer, y As Integer, w As Integer, h As Integer
 
      url = "https://logodix.com/logo/701195.jpg"
 
      x = Selection.Left
      y = Selection.Top
      w = Selection.Width
      h = Selection.Height
 
      ActiveSheet.Shapes.AddPicture url, msoFalse, msoTrue, x, y, w, h
 
  End Sub
```
