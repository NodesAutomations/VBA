
### Application.Wait
```vba
'makes the macro pause for approximately 10 seconds
Application.Wait (Now + TimeValue("0:00:10"))
```

### Run Macro At Specific Interval

```vba
Public interval As Double

Sub macro_timer()
    interval = Now + TimeValue("00:00:10")
    'Tells Excel when to next run the macro.
    Application.OnTime interval, "my_macro"
End Sub

Sub my_macro()
    'Macro code that you want to run.
    MsgBox "This is my sample macro output."
    'Calls the timer macro so it can be run again at the next interval.
    Call macro_timer
End Sub

Sub stop_macro()
    Application.OnTime earliesttime:=interval, procedure:="my_macro", schedule:=False
End Sub
```
