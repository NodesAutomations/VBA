VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Progress Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private Sub UserForm_Activate()
    code
End Sub

Sub code()
    Application.ScreenUpdating = False
    Dim i As Integer, j As Integer, pctCompl As Single

    Sheet1.Cells.Clear

    For i = 1 To 100
        For j = 1 To 1000
            Cells(i, 1).Value = j
        Next j
        pctCompl = i
        progress pctCompl
    Next i
    Application.ScreenUpdating = True
End Sub

Sub progress(pctCompl As Single)
    ProgressForm.Text.Caption = pctCompl & "% Completed"
    ProgressForm.Bar.Width = pctCompl * 2

    DoEvents
End Sub

