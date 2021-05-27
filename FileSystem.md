# Folders

### Create Folder
```vba
 If Len(Dir(FolderPath, vbDirectory)) = 0 Then
        MkDir FolderPath
 End If
```

# File

### Get FileName Without Extension

```vba
Public Sub Test2()
    Dim fso As New Scripting.FileSystemObject
    Debug.Print fso.GetBaseName("MyFile.something.txt")
End Sub
```

# Files

### Get List Of All FilePaths of Specific Files 

```vba
'Function to Get List of All PNG File Paths
Private Function GetListOfFilePaths() As Object
  
    'Get Folder Path of PNG DataBase
    Dim folderPath As String
    folderPath = ActivePresentation.Path & "\Data"
    
    'Set File System Object
    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    
    'Set SourceFoder
    Dim SourceFolder As Scripting.Folder
    Set SourceFolder = FSO.GetFolder(folderPath)
    
    'Get All PNG File Paths
    Dim FileItem As Scripting.File
    Set GetListOfFilePaths = CreateObject("System.Collections.ArrayList")
    
    For Each FileItem In SourceFolder.Files
        If Right(FileItem.Name, 4) = ".png" Then
            GetListOfFilePaths.Add FileItem.Path
        End If
    Next FileItem
    
End Function
```

