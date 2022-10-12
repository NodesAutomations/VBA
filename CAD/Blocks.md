### Loop through all Blocks in active drawing
```vba
Sub test()
    Dim cadApp As AcadApplication
    Set cadApp = GetObject(, "autocad.Application")
    
    Dim cadDoc As AcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As AcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    Dim entity As AcadEntity
    
    Dim ATTRIB_LIST  As Variant
    Dim attributeRef As AcadAttributeReference
    'Loop Through All Entity in Cad Model
    For Each entity In cadModel
        
        'Filter Block Entity
        If entity.ObjectName = "AcDbBlockReference" Then
            'Filter Specific Block using Name
            If entity.EffectiveName = "Pole" Then
                'Check if block contain Attributes
                If entity.HasAttributes = "True" Then
                    ATTRIB_LIST = entity.GetAttributes
                    Set attributeRef = ATTRIB_LIST(0)
                    Debug.Print attributeRef.TextString
                    attributeRef.TextString = "P1"
                End If
            End If
        End If
    Next
End Sub
```
Referance: [Get value block attributes from AutoCAD into excel with VBA](https://forums.autodesk.com/t5/vba/get-value-block-attributes-from-autocad-into-excel-with-vba/td-p/9446869)

### Code to Move and rotate Blocks
```vba

Public Sub Test()
    On Error GoTo ErrorHandler
    
    'Autocad Application
    Dim cadApp As AutoCAD.AcadApplication
    Set cadApp = GetObject(, "autocad.application")
    
    'Autocad Document
    Dim cadDoc As AutoCAD.AcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    'Autocad ModelSpace
    Dim cadModel As AutoCAD.AcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    'Loop through each entity
    Dim i As Integer
    Dim cadEntity As AutoCAD.AcadEntity
    Dim cadBlockRef As AutoCAD.AcadBlockReference
    
     Dim BasePoint(0 To 2) As Double
        'top left corner of text
    BasePoint(0) = 0: BasePoint(1) = 0: BasePoint(2) = 0
    
    For i = 0 To 0                               'cadModel.Count - 1
        'Convert Item to Cad Entity
        Set cadEntity = cadModel.Item(i)
        If cadEntity.ObjectName = "AcDbBlockReference" Then
            Set cadBlockRef = cadEntity
            cadBlockRef.InsertionPoint(0) = 0
            cadBlockRef.InsertionPoint(1) = 0
            'Get Insertion Point And rotation
            Debug.Print cadBlockRef.Name, cadBlockRef.InsertionPoint(0), cadBlockRef.InsertionPoint(1), cadBlockRef.Rotation
            cadBlockRef.Rotate cadBlockRef.InsertionPoint, 0.785
            cadBlockRef.Move cadBlockRef.InsertionPoint, BasePoint
        End If
    Next
    
    
    
Done:
    Exit Sub
ErrorHandler:
    If Err.Description <> "" Then
        MsgBox (Err.Description)
    End If
End Sub
```

### Code to Modify Dynamic Blocks Property
```vba

Public Sub Test()
    On Error GoTo ErrorHandler
    
    'Autocad Application
    Dim cadApp As AutoCAD.AcadApplication
    Set cadApp = GetObject(, "autocad.application")
    
    'Autocad Document
    Dim cadDoc As AutoCAD.AcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    'Autocad ModelSpace
    Dim cadModel As AutoCAD.AcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    'Loop through each entity
    Dim i As Integer, j As Integer
    Dim cadEntity As AutoCAD.AcadEntity
    Dim cadBlockRef As AutoCAD.AcadBlockReference
    Dim cadDynProps() As AutoCAD.AcadDynamicBlockReferenceProperty
    Dim var_atts As Variant
    
    Dim BasePoint(0 To 2) As Double
    'top left corner of text
    
    BasePoint(0) = 0: BasePoint(1) = 0: BasePoint(2) = 0
    
    For i = 0 To cadModel.Count - 1
        'Convert Item to Cad Entity
        Set cadEntity = cadModel.Item(i)
        If cadEntity.Handle = "436359" Then
            Set cadBlockRef = cadEntity
            Call dyn_prop(cadBlockRef, "Distance1", 1500)
            ''            If cadBlockRef.IsDynamicBlock Then
            ''
            ''                var_atts = cadBlockRef.GetDynamicBlockProperties
            ''
            ''                For j = LBound(var_atts) To UBound(var_atts)
            ''                    If var_atts(j).PropertyName = "Distance1" Then
            ''                        var_atts(j).Value = 1500
            ''                        Debug.Print var_atts(j).Value
            ''                    End If
            ''                Next
            ''            End If
        End If
    Next
    
Done:
    Exit Sub
ErrorHandler:
    If Err.Description <> "" Then
        MsgBox (Err.Description)
    End If
End Sub

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
Reference :[Get and Set Dynamic Blocks](https://forums.autodesk.com/t5/vba/get-and-set-dynamic-block-property/td-p/2977862)
