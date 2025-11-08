Attribute VB_Name = "Module1"
Option Explicit

Public file_type As String

Sub Create_Invoice()

Application.DisplayAlerts = False
Application.ScreenUpdating = False


Dim dsh As Worksheet
Dim tsh As Worksheet
Dim sh As Worksheet

Set dsh = ThisWorkbook.Sheets("Data")
Set tsh = ThisWorkbook.Sheets("Invoice Template")
Set sh = ThisWorkbook.Sheets("Support")

sh.Cells.Clear
dsh.AutoFilterMode = False

dsh.UsedRange.Copy sh.Range("A1")
sh.Range("D:Z").Clear
sh.UsedRange.RemoveDuplicates 1, xlYes

Dim i As Integer
Dim lr As Long
lr = Application.WorksheetFunction.CountA(dsh.Range("A:A"))

Dim folder_Path As String

folder_Path = ThisWorkbook.Sheets("Setting").Range("G6").Value

If Right(folder_Path, 1) <> "\" Then folder_Path = folder_Path & "\"

Dim nwb As Workbook

If Application.WorksheetFunction.CountA(sh.Range("A:A")) > 1 Then

    For i = 2 To Application.WorksheetFunction.CountA(sh.Range("A:A"))
        dsh.UsedRange.AutoFilter 1, sh.Cells(i, 1).Value
        tsh.Range("A12:D36").ClearContents
        dsh.Range("D2:G" & lr).SpecialCells(xlCellTypeVisible).Copy
        tsh.Range("A12").PasteSpecial xlPasteValuesAndNumberFormats
        
        tsh.Range("A7").Value = sh.Cells(i, 3).Value
        tsh.Range("D7").Value = sh.Cells(i, 1).Value
        tsh.Range("D8").Value = sh.Cells(i, 2).Value
        
        dsh.AutoFilterMode = False
        
        If file_type = "PDF" Then
            tsh.ExportAsFixedFormat xlTypePDF, folder_Path & sh.Cells(i, 1) & ".pdf"
        ElseIf file_type = "EXCEL" Then
            tsh.Copy
            Set nwb = ActiveWorkbook
            nwb.Sheets(1).UsedRange.Copy
            nwb.Sheets(1).UsedRange.PasteSpecial xlPasteValues
            nwb.SaveAs folder_Path & sh.Cells(i, 1) & ".xlsx"
            nwb.Close
        End If
        
    Next i

End If

MsgBox "Done"

End Sub

Sub Create_PDF_Invoice()

file_type = "PDF"
Call Create_Invoice
file_type = ""

End Sub


Sub Create_Excel_Invoice()

file_type = "EXCEL"
Call Create_Invoice
file_type = ""

End Sub

