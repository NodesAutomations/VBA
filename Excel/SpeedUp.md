### Tips
- Always Work with Excel Objects Directely don't use Select or Activate Methods or anyother workbook task
- Avoid Copy Paste Use Direct Assignments
- Turn On/Off Application Events like automatica calculation, screen Updates etc
- User Array Object for Storing Data because it's fastest object type
- Use Advance Filter If you're manipulating Ranges

Reference : [Macro Mastery Youtube](https://www.youtube.com/watch?v=GCSF5tq7pZ0)

### Code to Enable/Disable Excel Events
```vba
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub
Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub
```
Reference : [The SpreadSheet Guru](https://www.thespreadsheetguru.com/blog/2015/2/25/best-way-to-improve-vba-macro-performance-and-prevent-slow-code-execution)
