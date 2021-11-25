VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDiagnostics 
   ClientHeight    =   3510
   ClientLeft      =   15
   ClientTop       =   135
   ClientWidth     =   4590
   OleObjectBlob   =   "frmDiagnostics.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbWorksheets_Change()
Dim sheet   As Worksheet
Dim iCol    As Integer

    'clear current contents of available fields...
    lstFields.Clear
    
    'load fields from selected worksheet/table...
    Set sheet = Nothing
    Set sheet = ThisWorkbook.Worksheets(cmbWorksheets.Text)

    iCol = 1
    While (sheet.Cells(1, iCol) <> "")
    
        Call lstFields.AddItem(sheet.Cells(1, iCol))
    
        iCol = iCol + 1
    
    Wend
    
    Set sheet = Nothing
    
End Sub

Private Sub cmbWorksheets2_Change()
Dim sheet   As Worksheet
Dim iCol    As Integer

    'clear current contents of available fields...
    lstFields2.Clear
    
    'load fields from selected worksheet/table...
    Set sheet = Nothing
    Set sheet = ThisWorkbook.Worksheets(cmbWorksheets2.Text)

    iCol = 1
    While (sheet.Cells(1, iCol) <> "")
    
        Call lstFields2.AddItem(sheet.Cells(1, iCol))
    
        iCol = iCol + 1
    
    Wend
    
    Set sheet = Nothing
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdRunDiagnostics_Click()
Dim i As Integer
Dim iNumFields As Integer
Dim aFields() As String
Dim sFK1 As String
Dim sFK2 As String

    If cmbWorksheets.ListCount = 0 Then
        Exit Sub
    End If

    Call initDiagnosticSheet

    'gather selected fields...
    If lstDiagnostics.Selected(2) Then
    
        For i = 0 To lstFields.ListCount - 1
            If lstFields.Selected(i) Then
                sFK1 = lstFields.List(i)
                Exit For
            End If
        Next i
        
        For i = 0 To lstFields2.ListCount - 1
            If lstFields2.Selected(i) Then
                sFK2 = lstFields2.List(i)
                Exit For
            End If
        Next i
    
    Else
    
        ReDim aFields(0)
        iNumFields = -1
        For i = 0 To lstFields.ListCount - 1
    
            If lstFields.Selected(i) Then
                iNumFields = iNumFields + 1
                ReDim Preserve aFields(iNumFields)
                aFields(iNumFields) = lstFields.List(i)
            End If
    
        Next i
        
    End If

    
    If (lstDiagnostics.Selected(0)) Then
        'check for NULL diagnostic...
        Call checkForNull(cmbWorksheets.Text, aFields)
    ElseIf (lstDiagnostics.Selected(1)) Then
        'check for Duplicates diagnostic...
        Call checkForDuplicate(cmbWorksheets.Text, aFields)
    ElseIf (lstDiagnostics.Selected(2)) Then
        'cross check FKs...
        Call crossCheckFK(cmbWorksheets.Text, sFK1, cmbWorksheets2.Text, sFK2)
    End If
    
    Call formatDiagnosticSheet
    
    MsgBox "Diagnostic complete.", vbInformation, APPNAME
    
    'Unload Me
    
End Sub

Private Sub lstDiagnostics_Click()

    cmbWorksheets2.Enabled = lstDiagnostics.Selected(2)
    lstFields2.Enabled = lstDiagnostics.Selected(2)
    
    If lstDiagnostics.Selected(2) Then
        lstFields.MultiSelect = fmMultiSelectSingle
    Else
        lstFields.MultiSelect = fmMultiSelectMulti
    End If

End Sub

Private Sub UserForm_Activate()
Dim i As Integer

    Caption = APPNAME

    lstDiagnostics.Clear
    
    Call lstDiagnostics.AddItem("check for NULL")
    Call lstDiagnostics.AddItem("check for Duplicates")
    Call lstDiagnostics.AddItem("cross check FK")
    
    'default diagnostic...
    lstDiagnostics.ListIndex = 0
    cmbWorksheets2.Enabled = False
    lstFields2.Enabled = False

    'clear worksheets combo box...
    cmbWorksheets.Clear
    
    'load current/available worksheets/tables...
    For i = 1 To ThisWorkbook.Worksheets.Count
        
        If UCase(ThisWorkbook.Worksheets.Item(i).Name) <> "GENERAL" And UCase(ThisWorkbook.Worksheets.Item(i).Name) <> "DIAGNOSTIC" Then
        
            Call cmbWorksheets.AddItem(ThisWorkbook.Worksheets.Item(i).Name)
            Call cmbWorksheets2.AddItem(ThisWorkbook.Worksheets.Item(i).Name)
        
        End If
        
    Next i

End Sub
