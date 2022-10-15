Attribute VB_Name = "TableMacros"
Option Explicit

Sub FilterTable()
With Sheet1
If .Range("F3").Value = Empty Then Exit Sub
.Range("B7:J29").AdvancedFilter xlFilterInPlace, .Range("U3:U4"), , Unique:=False
.Shapes("ClearFiltBtn").Visible = msoCTrue
End With
BuildList
End Sub

Sub ClearFilter()
On Error Resume Next
Sheet1.ShowAllData
On Error GoTo 0
Sheet1.Range("7:29").EntireRow.Hidden = False
Sheet1.Range("F3").ClearContents
BuildList
End Sub

Sub DisplayPicture()
Dim SelRow As Long
With Sheet1
    On Error Resume Next
    .Shapes("ItemPic").Delete
    On Error GoTo 0
    SelRow = .Range("S3").Value 'Selected Row
    On Error Resume Next
    If Dir(.Range("J" & SelRow).Value) = "" Then Exit Sub
    With .Pictures.Insert(.Range("J" & SelRow).Value)
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Height = 100
            .Name = "ItemPic"
        End With
        End With
        With .Shapes("ItemPic")
        .Visible = msoCTrue
        .Left = Sheet1.Range("B" & SelRow + 1).Left
        .Top = Sheet1.Range("B" & SelRow + 1).Top
        End With
End With
End Sub
