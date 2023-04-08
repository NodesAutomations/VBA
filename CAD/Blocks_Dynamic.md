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

### Generalize Method to update dynamic property of block reference
```vba
  Set cadBlockRef = cadEntity
  Call dyn_prop(cadBlockRef, "Distance1", 1500)
```
```vba
Public Sub dyn_prop(objBlock As AcadBlockReference, name_of_property As String, value_of_property As Double)
    Dim i As Integer
    Dim dyn_properties() As AcadDynamicBlockReferenceProperty
    Dim var_atts As Variant

    var_atts = objBlock.GetDynamicBlockProperties

    For i = LBound(var_atts) To UBound(var_atts)
        If var_atts(i).PropertyName = name_of_property Then
            var_atts(i).Value = value_of_property
            ' ThisDrawing.SendCommand "_regen" & vbCr
        End If
    Next

End Sub
```
 
