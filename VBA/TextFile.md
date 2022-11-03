### Read Text File and Populate Array

```VBA
Private Sub ReadTextFileInArray()

    Dim FSO As Object, MyFile As Object
    Dim FileName As String, Arr As Variant

    FileName = "C:\Users\Intel7500\source\repos\Fiverr_muddsidekick_PictureRandomiser\Data.txt"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set MyFile = FSO.OpenTextFile(FileName, 1)
    Arr = Split(MyFile.ReadAll, vbNewLine)

End Sub
```
### Read Text File into single String
```vba
Private Sub ReadChatData()
    Dim chatFilePath As String
    chatFilePath = "C:\Users\Ryzen2600x\source\repos\Fiverr_muddsidekick_Contest\ChatLog.txt"
    
    Dim chatData As String
    Dim iFile As Integer: iFile = FreeFile
    Open chatFilePath For Input As #1
    chatData = Input(LOF(iFile), iFile)
    Close #1
    Debug.Print , chatData
End Sub
```
### Write Text file
  
```vba
Dim out As String
out = "D:\Documents\Excel sheets\VBA\output.txt"
Open out For Output As #1
'Add your Print code Here
Print #1,""
Close #1
```
