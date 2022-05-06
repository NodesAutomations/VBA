### Loop Through Entity

```vba
Sub Test()
    On Error GoTo ErrorHandler
    
    'Autocad Application
    Dim aCadApp As AutoCAD.aCadApplication
    Set aCadApp = GetObject(, "autocad.application")
    
    'Autocad Document
    Dim aCadDoc As AutoCAD.AcadDocument
    Set aCadDoc = aCadApp.ActiveDocument
    
    Dim entity As AutoCAD.AcadEntity
    Dim mtext As AutoCAD.AcadMText
    
    Sheet1.range("A1").CurrentRegion.ClearContents
    
    Dim rowId As Integer
    rowId = 1
    
    'Code to Loop Through all Items in AutoCAD
    For Each entity In aCadDoc.ModelSpace
        If entity.ObjectName = "AcDbMText" Then
            Set mtext = entity
            With Sheet1
                .Cells(rowId, 1) = mtext.InsertionPoint(0)
                .Cells(rowId, 2) = mtext.InsertionPoint(1)
                .Cells(rowId, 3) = mtext.TextString
            End With
            rowId = rowId + 1
        End If
    Next
        
    'Code to Loop Though User Selected Items
    Dim aCadSelectionSet As AutoCAD.aCadSelectionSet
    Set aCadSelectionSet = aCadDoc.SelectionSets.Add("Testddd")
    aCadSelectionSet.SelectOnScreen
    
    For Each entity In aCadSelectionSet
    
        If entity.ObjectName = "AcDbMText" Then
            Set mtext = entity
            With Sheet1
                .Cells(rowId, 1) = mtext.InsertionPoint(0)
                .Cells(rowId, 2) = mtext.InsertionPoint(1)
                .Cells(rowId, 3) = mtext.TextString
            End With
            rowId = rowId + 1
        End If
    Next
    
    aCadSelectionSet.Delete
    
    'Code to Loop Through Already Selected Items
    
    Dim aCadSelectionSet As AutoCAD.aCadSelectionSet
    Set aCadSelectionSet = aCadDoc.SelectionSets.Add(Now)
    aCadSelectionSet.Select acSelectionSetPrevious
    
    For Each entity In aCadSelectionSet
    
        If entity.ObjectName = "AcDbMText" Then
            Set mtext = entity
            With Sheet1
                .Cells(rowId, 1) = mtext.InsertionPoint(0)
                .Cells(rowId, 2) = mtext.InsertionPoint(1)
                .Cells(rowId, 3) = mtext.TextString
            End With
            rowId = rowId + 1
        End If
    Next
    
    '    aCadSelectionSet.Delete
    
Done:
    Exit Sub
ErrorHandler:
    If Err.Description <> "" Then
        MsgBox (Err.Description)
    End If

End Sub
```
