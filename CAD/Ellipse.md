### Sample Code to Draw Ellipse
```
Sub Test()
    Dim cadApp As ZcadApplication
    Set cadApp = GetObject(, "ZWcad.Application")
    
    Dim cadDoc As ZcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As ZcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    'Set Ellipse Parameter
    Dim majorRadius As Double
    Dim radiusRatio As Double
 
    majorRadius = 20
    radiusRatio = 0.75
    
    'Center Point Ellipse
    Dim centerPoint(0 To 2) As Double
    centerPoint(0) = 0: centerPoint(1) = 0#: centerPoint(2) = 0#

    
    'End Point of Major Axis
    'You can set angle of ellipse using this point
    Dim majorAxisEndPoint(0 To 2) As Double
    majorAxisEndPoint(0) = majorRadius#: majorAxisEndPoint(1) = 0#: majorAxisEndPoint(2) = 0#
    
    ' Add the ellipse to the model space
    Dim ellipseObj As Object
    Set ellipseObj = cadModel.AddEllipse(centerPoint, majorAxisEndPoint, radiusRatio)
    
End Sub
```

### Draw Partial Ellipse using start and end angle
```
Sub Test()
    Dim cadApp As ZcadApplication
    Set cadApp = GetObject(, "ZWcad.Application")
    
    Dim cadDoc As ZcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As ZcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    'Set Ellipse Parameter
    Dim majorRadius As Double
    Dim radiusRatio As Double
 
    majorRadius = 500
    radiusRatio = 0.45
    
    'Center Point Ellipse
    Dim centerPoint(0 To 2) As Double
    centerPoint(0) = 0: centerPoint(1) = 0#: centerPoint(2) = 0#

    
    'End Point of Major Axis
    'You can set angle of ellipse using this point
    Dim majorAxisEndPoint(0 To 2) As Double
    majorAxisEndPoint(0) = majorRadius#: majorAxisEndPoint(1) = 0#: majorAxisEndPoint(2) = 0#
    
    ' Add the ellipse to the model space
    Dim ellipseObj As Object
    Set ellipseObj = cadModel.AddEllipse(centerPoint, majorAxisEndPoint, radiusRatio)
    ellipseObj.StartAngle = 180 * (3.14159265358979 / 180)
    ellipseObj.EndAngle = 360 * (3.14159265358979 / 180)

End Sub
```
