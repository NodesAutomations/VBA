### Play Sound When Macro is Completed

```vba
Option Explicit
#If VBA7 Then
    '64 bit declares here
    Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As LongPtr
#Else
    Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
    '32 bit declares here
#End If

Sub Sound()
    Beep 880, 500
End Sub
```

### Audio Message
```vba
  Application.Speech.Speak "All Videos Are Success Fully Generated"
```
