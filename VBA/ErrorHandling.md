## Error Handling Syntax 

```vba
'Normal code with Runtime Error
Sub Test()
    Dim x As Integer
    x = "Test"
    Debug.Print x
End Sub

'Code with Defult Error Handling
'This Code will Behave Same as Normal Code
'Go to 0 mean jump to That Line
Sub Test_Default()
    On Error GoTo 0
    Dim x As Integer
    x = "Test"
    Debug.Print x
End Sub

'If You need to Ignore Erorr
Sub Test_Ignore_Error()
    On Error Resume Next
    Dim x As Integer
    x = "Test"
    Debug.Print x
End Sub

'Code with Error Handling
Sub Test_GoTO_Handler()
    On Error GoTo ErrorHandler
    Dim x As Integer
    x = "Test"
    Debug.Print x
    
Done:
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub
```
### Details of Error Object
![image](https://user-images.githubusercontent.com/60865708/126891528-7454754e-42c6-4b35-bc51-48c79bb1cf73.png)

### Error Handing Settings
- Default
    - ![image](https://user-images.githubusercontent.com/60865708/126891610-977365cd-4e04-4a38-b270-4a8708053cb7.png)

- Change To Break On All Error To Debug if Error Handing code is used and Set Back to Default
    - ![image](https://user-images.githubusercontent.com/60865708/126891604-761278ce-d913-4647-906d-5ad3d58965dd.png)

### References
- [Macro Mastery](https://excelmacromastery.com/vba-error-handling/)
- [Youtube Playlist : MacroMaster](https://www.youtube.com/playlist?list=PL7ScsebMq5uUc1sQaabDWcZuw9kvTYjis)

