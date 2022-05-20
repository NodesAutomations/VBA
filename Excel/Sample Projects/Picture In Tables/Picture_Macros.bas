Attribute VB_Name = "Picture_Macros"
Option Explicit
Dim PicPath As String
Dim SelRow As Long
Sub Pic_Browse_For_File()
Dim PicFile As FileDialog
Set PicFile = Application.FileDialog(msoFileDialogFilePicker)
With PicFile
    .Title = "Select Picture To Attach"
    .Filters.Add "All Picture Files", "*.jpg,*jpeg,*.gif,*.png,*.gif,*bmp,*.tiff", 1
    If .Show <> -1 Then GoTo NoSelection
    Sheet1.Range("L" & ActiveCell.Row).Value = .SelectedItems(1)  'Place Fill Filepath
End With
NoSelection:
End Sub
Sub Pic_Insert_Into_Cell()
With Sheet1
    SelRow = ActiveCell.Row
    If .Range("L" & SelRow).Value = Empty Then Pic_Browse_For_File 'Browse For Picture if non exist
    If .Range("L" & SelRow).Value = Empty Then GoTo NoPic 'Check for actual file path
    PicPath = .Range("L" & SelRow).Value
    On Error Resume Next
    .Shapes("Row" & SelRow & "Pic").Delete 'Delete any picture if it exists
    On Error GoTo 0
    With .Pictures.Insert(PicPath)
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Height = 50
            .Name = "Row" & SelRow & "Pic"
        End With
    End With
    With .Shapes("Row" & SelRow & "Pic")
        .Left = Sheet1.Range("K" & SelRow).Left
        .Top = Sheet1.Range("K" & SelRow).Top
        .IncrementLeft 8
        .IncrementTop 2
    End With
    .Range(SelRow & ":" & SelRow).RowHeight = 54
End With 'End With sheet 1
NoPic:
End Sub

Sub PIc_Instert_Into_Comment()
With Sheet1
    SelRow = ActiveCell.Row
    If .Range("L" & SelRow).Value = Empty Then Pic_Browse_For_File 'Browse For Picture if non exist
    PicPath = .Range("L" & SelRow).Value
    On Error Resume Next 'Catch bug if no existing comment
    .Range("E" & SelRow).Comment.Delete 'Delete any existing comment first
    On Error GoTo 0
    .Range("E" & SelRow).AddComment
    With .Range("E" & SelRow).Comment
        .Visible = True
        .Shape.Fill.UserPicture PicPath
        .Text ""
        .Shape.Select True
        Selection.ShapeRange.ScaleWidth 0.65, msoFalse, msoScaleFromTopLeft
        .Visible = False
    End With
End With
End Sub

Sub Pic_Display_On_Click()
With Sheet1
SelRow = ActiveCell.Row
If .Range("L" & SelRow).Value = Empty Then Pic_Browse_For_File 'Browse For Picture if non exist
If .Range("L" & SelRow).Value = Empty Then GoTo NoPic 'Check for actual file path
PicPath = .Range("L" & SelRow).Value
On Error Resume Next
.Shapes("SelectionRowPic").Delete 'Delete any picture if it exists
On Error GoTo 0
With .Pictures.Insert(PicPath)
    With .ShapeRange
        .LockAspectRatio = msoTrue
        .Height = 150
        .Name = "SelectionRowPic"
    End With
End With
With .Shapes("SelectionRowPic")
        .Left = Sheet1.Range("M" & SelRow).Left
        .Top = Sheet1.Range("M" & SelRow).Top
End With
End With
NoPic:
End Sub

Sub Pic_DisplayAllPics()
Dim LastRow, TabRow As Long
LastRow = Sheet1.Range("E9999").End(xlUp).Row
For TabRow = 5 To LastRow
    Sheet1.Range("E" & TabRow).Select
    Pic_Insert_Into_Cell
Next TabRow
End Sub
