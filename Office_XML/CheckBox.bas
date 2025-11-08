Attribute VB_Name = "CheckBox"
'@Folder("VBAProject")
Option Explicit

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Dim ribRibbon As IRibbonUI
    Set ribRibbon = ribbon
    ribRibbon.ActivateTab ("CustomTab")
End Sub

'Callback for mycheckbox onAction
Sub CheckBox_OnAction(control As IRibbonControl, pressed As Boolean)
    If pressed Then
        Sheet1.Range("A1") = 1
    Else
        Sheet1.Range("A1") = 0
    End If
End Sub

Public Sub CheckBox_OnGetPressed( _
       ByRef control As IRibbonControl, _
       ByRef pressed As Variant)

    If control.ID = "mycheckbox" Then
        pressed = CBool(Sheet1.Range("A1"))
    End If
End Sub


