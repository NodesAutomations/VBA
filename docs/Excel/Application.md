### Application Caller
- Returns information about how Visual Basic was called
- We can use it with cells, buttons,Shapes
```vba
Sub Shape_Click()
Dim Sh As Shape
Set Sh = ActiveSheet.Shapes(Application.Caller)
MsgBox Sh.Name
End Sub
```
### Application Event

```vba
'The WithEvents keyword makes the appevent variable available in the Object drop-down in the Class1 (Code) module window.
Public WithEvents appevent As Application
Private Sub appevent_WindowResize(ByVal Wb As Excel.Workbook,ByVal Wn As Excel.Window)
       
End Sub
```
- [Microsoft Ref](https://docs.microsoft.com/en-us/office/troubleshoot/excel/create-application-level-event-handler)
- [Wise Owl](https://www.youtube.com/watch?v=SaQfOIeOuHk)
