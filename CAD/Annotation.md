# Add Text
```vba
Public Function addMtext(ByVal originX As Double, ByVal originY As Double, ByVal width As Double, ByVal height As Double, ByVal text As String) As Object
    Dim corner(0 To 2) As Double
    'top left corner of text
    corner(0) = originX: corner(1) = originY: corner(2) = 0

    Set addMtext = cadModel.addMtext(corner, width, text)
    addMtext.height = height

End Function

Sub test()
    Dim cadApp As ZcadApplication
    Set cadApp = GetObject(, "ZWcad.Application")
    
    Dim cadDoc As ZcadDocument
    Set cadDoc = cadApp.ActiveDocument
    
    Dim cadModel As ZcadModelSpace
    Set cadModel = cadDoc.ModelSpace
    
    
    'Add Text
    Dim corner(0 To 2) As Double
    'top left corner of text
    corner(0) = 0: corner(1) = 0: corner(2) = 0

    Dim cadMtext As ZcadMText
    Set cadMtext = cadModel.addMtext(corner, 20, "{\LBore Hole BH-1 A1 RHS}")
    cadMtext.height = 400
    cadMtext.Color = CadColors.Magenta
    cadMtext.StyleName = "ROMANT"
    cadMtext.AttachmentPoint = zcAttachmentPointMiddleCenter
    Debug.Print cadMtext.StyleName
    cadMtext.Update
    cadDoc.Regen zcActiveViewport
    cadDoc.Application.ZoomAll

End Sub

```
