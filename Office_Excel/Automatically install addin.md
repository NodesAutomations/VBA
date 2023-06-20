### Create Excel sheet to install addin
- create `install.xlsm` with module containing below code
- you can create button to run this or add workbook_open event
```

Sub install_add_in()

Dim mypath As String, strfile As String, fileName As String

mypath = ActiveWorkbook.Path & "\"
fileName = "Test_Addin"   'replace General.Tools with the name of your Add-in !!!
strfile = "" & fileName & ".xlam"

file_to_copy = mypath & strfile

folder_to_copy = Environ("Appdata") & "MicrosoftAddIns"

copied_file = folder_to_copy & strfile

'Check if add-in is installed
If Len(Dir(copied_file)) = 0 Then

'if add-in does not exist then copy the file
FileCopy file_to_copy, copied_file
AddIns(fileName).Installed = True
MsgBox "Add-in installed"

Else

'if add-in already exists then the user will decide if will replace it or not
x = MsgBox("Add-in allready exists ! Replace ?", vbYesNo)

    If x = vbNo Then
        Exit Sub
    ElseIf x = vbYes Then
        
        'deactivate the add-in if it is activated
        If AddIns(fileName).Installed = True Then
            AddIns(fileName).Installed = False
        End If
        
        'delete the old file
        Kill copied_file
        
        'copy the new file
        FileCopy file_to_copy, copied_file
        AddIns(fileName).Installed = True
        MsgBox "New Add-in Installed !"

    End If

End If

End Sub
```
