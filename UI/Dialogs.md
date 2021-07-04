### Message Box

```VBA
Function MessageBox_Demo() 
   'Message Box with just prompt message 
   MsgBox("Welcome")     
   
   'Message Box with title, yes no and cancel Butttons  
   int a = MsgBox("Do you like blue color?",3,"Choose options") 
   ' Assume that you press No Button  
   msgbox ("The Value of a is " & a) 
End Function
```
Excel VBA MsgBox Icon Constants
- vbCritical	Shows the critical message icon
- vbQuestion	Shows the question icon
- vbExclamation	Shows the warning message icon
- vbInformation	Shows the information icon

```vba
Sub Test()
    Dim result As Integer
    result = MsgBox("Running this Macro will Delete All text From that Shape", vbOKCancel + vbCritical, "Shape Contain Text")
    If result = vbCancel Then
    End If
End Sub
```
Additional Ref for MesageBox
- [VBA MessageBox](https://trumpexcel.com/vba-msgbox/)
