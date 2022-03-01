Option Explicit

'Ref: https://www.contextures.com/xldataval10.html
'==========================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim str As String
    Dim cboTemp As OLEObject
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Set cboTemp = ws.OLEObjects("TempCombo")
    On Error Resume Next
    With cboTemp
        'clear and hide the combo box
        .ListFillRange = ""
        .LinkedCell = ""
        .Visible = False
    End With
    On Error GoTo errHandler
    If Target.Validation.Type = 3 Then
        'if the cell contains a data validation list
        Cancel = True
        Application.EnableEvents = False
        'get the data validation formula
        str = Target.Validation.Formula1
        str = Right(str, Len(str) - 1)
        With cboTemp
            'show the combobox with the list
            .Visible = True
            .Left = Target.Left
            .Top = Target.Top
            .Width = Target.Width + 5
            .Height = Target.Height + 5
            .ListFillRange = ws.Range(str).Address
            .LinkedCell = Target.Address
        End With
        cboTemp.Activate
        'open the drop down list automatically
        Me.TempCombo.DropDown
    End If
  
errHandler:
    Application.EnableEvents = True
    Exit Sub

End Sub

'=========================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim str As String
    Dim cboTemp As OLEObject
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Set cboTemp = ws.OLEObjects("TempCombo")
    On Error Resume Next
    If cboTemp.Visible = True Then
        With cboTemp
            .Top = 10
            .Left = 10
            .ListFillRange = ""
            .LinkedCell = ""
            .Visible = False
            .Value = ""
        End With
    End If

errHandler:
    Application.EnableEvents = True
    Exit Sub

End Sub

'====================================
'Optional code to move to next cell if Tab or Enter are pressed
'from code by Ted Lanham
'***NOTE: if KeyDown causes problems, change to KeyUp

Private Sub TempCombo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
    Case 9                                       'Tab
        ActiveCell.Offset(0, 1).Activate
    Case 13                                      'Enter
        ActiveCell.Offset(1, 0).Activate
    Case Else
        'do nothing
    End Select
End Sub

'====================================



