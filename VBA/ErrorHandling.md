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
![image](https://user-images.githubusercontent.com/60865708/126891528-7454754e-42c6-4b35-bc51-48c79bb1cf73.png)

### References
- [Youtube Playlist : MacroMaster](https://www.youtube.com/playlist?list=PL7ScsebMq5uUc1sQaabDWcZuw9kvTYjis)

