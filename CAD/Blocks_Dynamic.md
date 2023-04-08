### Code to Update Dynamimc Block Properties
```vba
Sub test()
    Dim cadApp As AcadApplication
    Set cadApp = GetObject(, "autocad.Application")
    
    Dim cadDoc As AcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As AcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    'Loop through each entity
    Dim i As Integer
    Dim cadEntity As AutoCAD.AcadEntity
    Dim cadBlockRef As AutoCAD.AcadBlockReference
    Dim prop As Variant
    For Each cadEntity In cadModel
        'Filter Block Entity
        If cadEntity.ObjectName = "AcDbBlockReference" Then
            'Filter Specific Block using Name
            If cadEntity.EffectiveName = "Test" Then
                Set cadBlockRef = cadEntity
                If cadBlockRef.IsDynamicBlock Then
                    prop = cadBlockRef.GetDynamicBlockProperties
                    For i = LBound(prop) To UBound(prop)
                        If prop(i).PropertyName = "HEIGHT" Then
                            prop(i).Value = 500#
                        End If
                    Next
                End If
            End If
        End If
    Next
End Sub
```
