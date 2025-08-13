### Sample code to draw hatch using internal point
```vba
Sub AddHatch()
    Dim cadApp As ZcadApplication
    Set cadApp = GetObject(, "ZWcad.Application")
    
    Dim cadDoc As ZcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As ZcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    Dim rectangle As ZcadPolyline
    Set rectangle = AddRectangle(cadModel, 0, 0, 100, 100)
    
    'Create an Outer Boundary polyline at the selected point
    Dim insertionPnt As String
    insertionPnt = 50 & "," & 50 & ",0"
    cadDoc.SendCommand "-Boundary" & vbCr & insertionPnt & vbCr & vbCr
 
 
    'Make an array with last created entity (Boundary Polyline)
    Dim arr(0 To 0) As ZcadEntity
    Set arr(0) = cadDoc.ModelSpace.Item(cadDoc.ModelSpace.Count - 1)
    
    Dim patternname As String
    Dim patterntype As Long
    Dim bassociativity As Boolean
    
    patternname = "SOLID"
    patterntype = zcHatchPatternTypePreDefined
    bassociativity = True
    
    Dim hatchObj As ZcadHatch
    Set hatchObj = cadModel.AddHatch(patterntype, patternname, bassociativity)
    hatchObj.PatternScale = 1
    hatchObj.AppendOuterLoop (arr)
    hatchObj.Evaluate
End Sub

Public Function AddRectangle(ByRef cadModel As ZcadModelSpace, ByVal originX As Double, ByVal originY As Double, ByVal width As Double, ByVal height As Double) As Object
    Dim coords(11) As Double
    coords(0) = originX: coords(1) = originY: coords(2) = 0
    coords(3) = coords(0) + width: coords(4) = coords(1): coords(5) = 0
    coords(6) = coords(3): coords(7) = coords(4) + height: coords(8) = 0
    coords(9) = coords(6) - width: coords(10) = coords(7): coords(11) = 0
    
    Set AddRectangle = cadModel.AddPolyline(coords)
    AddRectangle.Closed = True
End Function

```