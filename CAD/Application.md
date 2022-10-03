# Start Up

### AutoCAD
```vba
Sub test()
    Dim cadApp As AcadApplication
    Set cadApp = GetObject(, "autocad.Application")
    
    Dim cadDoc As AcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As AcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    
    'Add Text
    Dim corner(0 To 2) As Double
    'top left corner of text
    corner(0) = 0: corner(1) = 0: corner(2) = 0

    Dim addMtext As Object
    Set addMtext = cadModel.addMtext(corner, 20, "Test")
    addMtext.Height = Height
End Sub
```
### AutoCAD
```vba
Sub test()
    Dim cadApp As ZcadApplication
    Set cadApp = GetObject(, "ZWcad.Application")
    
    Dim cadDoc As ZcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As ZcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    
    'Add Text
    Dim corner(0 To 2) As Double
    'top left corner of text
    corner(0) = 0: corner(1) = 0: corner(2) = 0

    Dim addMtext As Object
    Set addMtext = cadModel.addMtext(corner, 20, "Test")
    addMtext.height = height
End Sub
```
