### Tips
- Always Work with Excel Objects Directely don't use Select or Activate Methods or anyother workbook task
- 

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
