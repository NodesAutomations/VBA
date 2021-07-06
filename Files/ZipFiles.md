### UnZip File
First Include reference to Microsoft Shell Controls and Automation
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
