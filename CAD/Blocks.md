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
