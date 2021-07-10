# Common

### Code To Loop Through All Files in Folder
```vba
Public Sub FindAllPPTFiles()
    Dim fso As FileSystemObject
    Dim libraryPath As String

    libraryPath = ThisWorkbook.CustomDocumentProperties("LibraryPath")

    Set fso = New FileSystemObject
    DoFolder fso.GetFolder(libraryPath)
End Sub

Sub DoFolder(Folder)
    Dim subFolder As Folder
    For Each subFolder In Folder.SubFolders
        DoFolder subFolder
    Next
    Dim file As file
    For Each file In Folder.Files
        Debug.Print , file.Path
    Next
End Sub
```


# Folders
### Check if Folder Exist
```vba
Sub CheckFolderExists ()

Dim strFolderName As String
Dim strFolderExists As String

    strFolderName = "C:\Users\Nikola\Desktop\VBA articles\Test Folder\"
    strFolderExists = Dir(strFolderName, vbDirectory)

    If strFolderExists = "" Then
        MsgBox "The selected folder doesn't exist"
    Else
        MsgBox "The selected folder exists"
    End If

End Sub
```

### Create Folder
```vba
 If Len(Dir(FolderPath, vbDirectory)) = 0 Then
        MkDir FolderPath
 End If
```

### FilePath When Saving it in OneDrive
```vba
Sub TestLocalFullName()
    Debug.Print "URL: " & ActiveWorkbook.FullName
    Debug.Print "Local: " & LocalFullName(ActiveWorkbook.FullName)
    Debug.Print "Test: " & Dir(LocalFullName(ActiveWorkbook.FullName))
End Sub

Private Function LocalFullName$(ByVal fullPath$)
    'Finds local path for a OneDrive file URL, using environment variables of OneDrive
    'Reference https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    'Authors: Philip Swannell 2019-01-14, MatChrupczalski 2019-05-19, Horoman 2020-03-29, P.G.Schild 2020-04-02

    Dim ii&
    Dim iPos&
    Dim oneDrivePath$
    Dim endFilePath$

    If Left(fullPath, 8) = "https://" Then 'Possibly a OneDrive URL
        If InStr(1, fullPath, "my.sharepoint.com") <> 0 Then 'Commercial OneDrive
            'For commercial OneDrive, path looks like "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
            'Find "/Documents" in string and replace everything before the end with OneDrive local path
            iPos = InStr(1, fullPath, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without pointer in OneDrive. Include leading "/"
        Else 'Personal OneDrive
            'For personal OneDrive, path looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
            'We can get local file path by replacing "https.." up to the 4th slash, with the OneDrive local path obtained from registry
            iPos = 8 'Last slash in https://
            For ii = 1 To 2
                iPos = InStr(iPos + 1, fullPath, "/") 'find 4th slash
            Next ii
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without OneDrive root. Include leading "/"
        End If
        endFilePath = Replace(endFilePath, "/", Application.PathSeparator) 'Replace forward slashes with back slashes (URL type to Windows type)
        For ii = 1 To 3 'Loop to see if the tentative LocalWorkbookName is the name of a file that actually exists, if so return the name
            oneDrivePath = Environ(Choose(ii, "OneDriveCommercial", "OneDriveConsumer", "OneDrive")) 'Check possible local paths. "OneDrive" should be the last one
            If 0 < Len(oneDrivePath) Then
                LocalFullName = oneDrivePath & endFilePath
                Exit Function 'Success (i.e. found the correct Environ parameter)
            End If
        Next ii
        'Possibly raise an error here when attempt to convert to a local file name fails - e.g. for "shared with me" files
        LocalFullName = vbNullString
    Else
        LocalFullName = fullPath
    End If
End Function

```
# File

### Check if File Exist
```vba
Sub CheckFileExists ()

Dim strFileName As String
Dim strFileExists As String

    strFileName = "C:\Users\Nikola\Desktop\VBA articles\Test File Exists.xlsx"
    strFileExists = Dir(strFileName)

   If strFileExists = "" Then
        MsgBox "The selected file doesn't exist"
    Else
        MsgBox "The selected file exists"
    End If

End Sub
```
### Get FileName Without Extension

```vba
Public Sub Test2()
    Dim fso As New Scripting.FileSystemObject
    Debug.Print fso.GetBaseName("MyFile.something.txt")
End Sub
```
### Extract File Data Using FileSystem Object
Add a reference to Microsoft Scripting Runtime
Early Binding
```vba
Dim fso as new FileSystemObject
Dim fileName As String
fileName = fso.GetFileName("c:\any path\file.txt")
```
Late Binding
```vba
With CreateObject("Scripting.FileSystemObject")
    fileName = .GetFileName(FilePath)
    extName = .GetExtensionName(FilePath)
    baseName = .GetBaseName(FilePath)
    parentName = .GetParentFolderName(FilePath)
End With
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

