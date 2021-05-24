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
