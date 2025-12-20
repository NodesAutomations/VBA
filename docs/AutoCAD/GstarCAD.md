# GStarCad

## Sample VBA

```visualbasic
Sub CreateTable()
    'Get AutoCad App
    Dim cadApp As Object
    Set cadApp = GetObject(, "gcad.Application")
    
    'Get active AutoCAD Drawing
    Dim cadDoc As Object
    Set cadDoc = cadApp.ActiveDocument
    
    'Get model space
    Dim cadModel As Object
    Set cadModel = cadDoc.ModelSpace
    
    'Using 0,0 as table top left base point
    Dim basePoint(0 To 2) As Double
    basePoint(0) = 0: basePoint(1) = 0: basePoint(2) = 0
    
    'Create AutoCAD Table
    Dim table As Object
    Set table = cadDoc.ModelSpace.AddTable(basePoint, 4, 3, 0.6, 2.4)
 
    With table
        'Unmerge Header row
        .UnmergeCells 0, 0, 0, 2
        
        'Header Row
        .SetText 0, 0, "BARID"
        .SetText 0, 1, "DIA"
        .SetText 0, 2, "LENGTH"
        
        'Row 1
        .SetText 1, 0, "1"
        .SetText 1, 1, "10"
        .SetText 1, 2, "5"
     
        'Row 2
        .SetText 2, 0, "2"
        .SetText 2, 1, "12"
        .SetText 2, 2, "10"
        
        'Row 3
        .SetText 3, 0, "3"
        .SetText 3, 1, "16"
        .SetText 3, 2, "15"
    End With
    
End Sub
```
