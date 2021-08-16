### Basic Access Modifier
```vba
Option Explicit

'Available thorough all Modules
'Exposed as Macro
Sub Test()
    MsgBox "Hello"
End Sub

'Available thorough all Modules
'Exposed as Macro
Public Sub PublicTest()
    MsgBox "Hello"
End Sub

'Only Available in that Module
Private Sub PrivateTest()
    MsgBox "Hello"
End Sub
```
Same Modifier Rules Apply for Function also
Also We can call public Methods directly without using call

### Protip
- Use `Option Private Module` In Beginning/ Top of Module to Stop being exposed as macro or UDF
