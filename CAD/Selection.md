### Code to Loop through Selected objects

```vba
Public Sub ExtractBlockData()
    'Figure out Way to Get Selected Items from AutoCad

    'Get Active AutoCad Application
    Dim acadApp As AcadApplication
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    
    'If AutoCAD is not opened create a new instance and make it visible.
    If acadApp Is Nothing Then
        MsgBox "AutoCad Is not open"
        Exit Sub
    End If

    'Get Active Autocad Drawing
    Dim acadDoc As AcadDocument
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo 0

    'No active drawing found. Create a new one.
    If acadDoc Is Nothing Then
        MsgBox "No Active AutoCad File"
        Exit Sub
    End If
     
    Dim selectionSet As AcadSelectionSet
    Set selectionSet = acadDoc.PickfirstSelectionSet
    
    'If nothing is selected just exit
    If selectionSet.Count = 0 Then
        Exit Sub
    End If
 
    Dim i As Integer
    Dim cadEntity As autocad.AcadEntity
    Dim cadBlockRef As autocad.AcadBlockReference
    Dim cadAttributeRef As autocad.AcadAttributeReference
    Dim ATTRIB_LIST  As Variant
     
    For Each cadEntity In selectionSet
        If cadEntity.ObjectName = "AcDbBlockReference" Then
            Set cadBlockRef = cadEntity
            'Check if block contain Attributes
            If cadBlockRef.HasAttributes = "True" Then
                ATTRIB_LIST = cadBlockRef.GetAttributes
                For i = LBound(ATTRIB_LIST) To UBound(ATTRIB_LIST)
                    Set cadAttributeRef = ATTRIB_LIST(i)
                    Debug.Print cadAttributeRef.TagString
                Next
            End If
        End If
    Next
End Sub
```
