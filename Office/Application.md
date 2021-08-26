### Application Caller
```vba

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
