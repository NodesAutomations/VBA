### UnZip File
This Code uses the Windows default file compressor.
- reference 
    - Microsoft Shell Controls and Automation
```vba
Sub Unzip()
    'Define Variable Data Types
    Dim zipFileName As String
    Dim unZipFolderName As String
    Dim objZipItems As FolderItems
    Dim objZipItem As FolderItem
    
    'Set Zip File Name & Folder path to Unzip
    zipFileName = "C:\Users\Ryzen2600x\Downloads\New Microsoft PowerPoint Presentation.zip"
    unZipFolderName = "C:\Users\Ryzen2600x\Downloads"
    
    'Early Binding Reference
    'Add Tools -> Reference -> "Microsoft Shell Controls & Automation"
    Dim wShApp As Shell
    Set wShApp = CreateObject("Shell.Application")
    Set objZipItems = wShApp.Namespace(zipFileName).items
    
    'Extract: Unzip all Files to Folder
    wShApp.Namespace(unZipFolderName).CopyHere objZipItems
    
End Sub
```
### Zip File
- reference 
    - Microsoft Shell Controls and Automation
```vba
Sub ZipVBA()
    'Define Variable Data Types
    'Early Binding Reference
    'Add Tools -> Reference -> "Microsoft Shell Controls & Automation"
    Dim zipFileName As String
    Dim unZipFolderName As String
    Dim objZipItems As FolderItems
    Dim objZipItem As FolderItem
    
    'Set Zip File Name & Folder path to Unzip
  zipFileName = "C:\Users\Ryzen2600x\Downloads\New Microsoft PowerPoint Presentation.zip"
    unZipFolderName = "C:\Users\Ryzen2600x\Downloads\New Microsoft PowerPoint Presentation"
    
    'Create Empty Zip file
    Open zipFileName For Output As #1
        Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Initialize Shell Object & File to be Zipped
    Dim wShApp As Shell
    Set wShApp = CreateObject("Shell.Application")
    Set objZipItems = wShApp.Namespace(unZipFolderName).Items
    
    'Method1: Compress All Files at once
    wShApp.Namespace(zipFileName).CopyHere objZipItems
    Do Until wShApp.Namespace(zipFileName).Items.Count = objZipItems.Count
        DoEvents
        Debug.Print "Processing " & wShApp.Namespace(zipFileName).Items.Count & " of" & objZipItems.Count
        Application.Wait DateAdd("s", 1, Now)
    Loop
    Debug.Print "Processing " & wShApp.Namespace(zipFileName).Items.Count & " of" & objZipItems.Count
    
    'Method2: ZIP Files One by one
    For Each objZipItem In objZipItems
        If objZipItem.Name = "FileName.ext" Then
            wShApp.Namespace(zipFileName).CopyHere objZipItem
        End If
    Next
    
End Sub
```
