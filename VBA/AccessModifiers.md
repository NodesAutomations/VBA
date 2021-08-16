### Basic Access Modifier
```vba
Option Explicit

'Available thorough Macro
Sub Test()
    MsgBox "Hello"
End Sub

'Available thorough Macro
Public Sub PublicTest()
    MsgBox "Hello"
End Sub

'Only Available in that Module
Private Sub PrivateTest()
    MsgBox "Hello"
End Sub

```
