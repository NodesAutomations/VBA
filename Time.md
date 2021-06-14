### VBA Function TO Delay Execution In Milliseconds
```vba
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        'To delay program execution for 250 milliseconds:
        Delay 500
        Debug.Print TimeValue(Now)
    Next
End Sub
```
```vba
'VBA function to delay execution:

Function Delay(ms)
    Delay = Timer + ms / 1000
    While Timer < Delay: DoEvents: Wend
End Function

'To delay program execution for 250 milliseconds:
Delay 250
```
```vba
'Excel VBA includes the Wait method on the Application object:
Application.Wait Now + TimeValue("00:00:25")
'The above will delay VBA execution for 25 seconds. But 
'Application.Wait is unreliable for delays less than a second.

'There is also the Sleep Win32 API, but the Delay() function above 
'works better. But here is how to use Sleep. Place the following
'declarations at the top of a code module:

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'Now call like so:

Sleep 250		'this causes Excel and VBA to go dormant for 250 ms.


'Both Application.Wait and Sleep() block the entire Excel interface. 
'Excel literally goes to sleep until the wait period is over. Delay() 
'does not block the Excel interface; it just prevents VBA from 
'continuing until the delay is over.
```

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
