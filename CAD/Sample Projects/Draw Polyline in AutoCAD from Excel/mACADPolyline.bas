Attribute VB_Name = "mACADPolyline"
Option Explicit
Option Private Module

Sub DrawPolyline()

    'Draws a polyline in AutoCAD using X and Y coordinates from sheet Coordinates.
 
    'In order to use the macro you must enable the AutoCAD library from VBA editor:
    'Go to Tools -> References -> Autocad xxxx Type Library, where xxxx depends
    'on your AutoCAD version (i.e. 2010, 2011, 2012, etc.) you have installed to your PC.
        
    'Declaring the necessary variables.
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim LastRow As Long
    Dim acadPol As AcadLWPolyline
    Dim dblCoordinates() As Double
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    shCoordinates.Activate
    
    'Find the last row.
    With shCoordinates
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
        
    'Check if there are at least two points.
    If LastRow < 3 Then
        MsgBox "There not enough points to draw the polyline!", vbCritical, "Points Error"
        Exit Sub
    End If
        
    'Check if AutoCAD is open.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    
    'If AutoCAD is not opened create a new instance and make it visible.
    If acadApp Is Nothing Then
        Set acadApp = New AcadApplication
        acadApp.Visible = True
    End If
    
    'Check if there is an active drawing.
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo 0
    
    'No active drawing found. Create a new one.
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
        acadApp.Visible = True
    End If
    
    'Get the array size.
    ReDim dblCoordinates(2 * (LastRow - 1) - 1)
    
    'Pass the coordinates to array.
    k = 0
    For i = 2 To LastRow
        For j = 1 To 2
            dblCoordinates(k) = shCoordinates.Cells(i, j)
            k = k + 1
        Next j
    Next i
    
    'Draw the polyline either at model space or at paper space.
    If acadDoc.ActiveSpace = acModelSpace Then
        Set acadPol = acadDoc.ModelSpace.AddLightWeightPolyline(dblCoordinates)
    Else
        Set acadPol = acadDoc.PaperSpace.AddLightWeightPolyline(dblCoordinates)
    End If
    
    'Leave the polyline open (the last point is not connected with the first point.
    'Set the next line to true if you need to connect the last point with the first one.
    acadPol.Closed = False
    acadPol.Update
    
    'Zooming in to the drawing area.
    acadApp.ZoomExtents
    
    'Inform the user that the polyline was created.
    MsgBox "The polyline was successfully created!", vbInformation, "Finished"

End Sub

Sub ClearCoordinates()
    
    Dim LastRow As Long
    
    shCoordinates.Activate
    
    'Find the last row.
    With shCoordinates
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:B" & LastRow).ClearContents
        .Range("A2").Select
    End With
    
End Sub
