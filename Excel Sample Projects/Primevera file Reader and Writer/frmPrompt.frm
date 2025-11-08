VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrompt 
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   OleObjectBlob   =   "frmPrompt.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
Dim i As Integer
Dim dlgType As MsoFileDialogType
Dim dlg As FileDialog
Dim fdfs As FileDialogFilters
Dim fdf As FileDialogFilter

    'initialize working folder...
    If sWorkingFolder = "" Then
        sWorkingFolder = "c:\"
    End If
    
    


If bMode Then
    dlgType = msoFileDialogOpen
Else
    dlgType = msoFileDialogSaveAs
End If

        Set dlg = Application.FileDialog(dlgType)

    With dlg
        If bMode Then
            .Filters.Clear
            .Filters.Add "XER", "*.xer; *.xer"
            .InitialFileName = ""
        Else
            .InitialFileName = sWorkingFolder & "\*.xer"
        End If
        '.lpstrInitialDir = sWorkingFolder
        .Title = APPNAME
        .Show

        If .SelectedItems.Count > 0 Then
            If Strings.Trim(.SelectedItems.Item(1)) <> "" Then

                If bMode Then
                    txtXerFile.Text = .SelectedItems(1)
                Else
                    txtXerFile.Text = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), ".") - 1)
                End If
                
                'set new working folder...
                For i = Len(.SelectedItems(1)) To 1 Step -1
    
                    If Strings.Mid(.SelectedItems(1), i, 1) = "\" Then
                        sWorkingFolder = Strings.Mid(.SelectedItems(1), 1, i - 1)
                        Exit For
                    End If
    
                Next i
    
            End If
        End If
    End With

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If (Strings.Trim(txtXerFile.Text) <> "") Then
    
        sXerFileName = txtXerFile.Text
        Worksheets("General").Cells(3, 1) = "XER File:"
        Worksheets("General").Cells(3, 1).Font.Bold = True
        Worksheets("General").Cells(3, 1).Font.Color = RGB(150, 50, 150)
        Worksheets("General").Cells(4, 1) = sXerFileName
        Worksheets("General").Cells(4, 1).Font.Bold = True
        
        If (frmPrompt.chkLoadStatsOnly.Value = True) Then
            Worksheets("General").Cells(5, 1) = "Statistics only"
            Worksheets("General").Cells(5, 1).Font.Italic = True
        End If
    
    Else
    
        MsgBox "XER file not specified.", vbExclamation, APPNAME
        Exit Sub
        
    End If
    
    Unload Me

End Sub

Private Sub UserForm_Activate()

    Caption = APPNAME

End Sub
