### Using Environ function
```vba
environ("Path") 
```


### Using CMD
```vba
Sub RunCMD()
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0 'Change this to 0 to hide the window
    Dim cmd As String
    Dim output As String
    'Change the cmd variable to your desired command
    cmd = "where python"
    'Run the command and get the output
    output = wsh.Exec("cmd.exe /S /C " & cmd).StdOut.ReadAll
    
    Dim paths As Variant
    paths = Split(output, vbNewLine)
    
    'Display the output in a message box
    Debug.Print paths(0)
End Sub
```
