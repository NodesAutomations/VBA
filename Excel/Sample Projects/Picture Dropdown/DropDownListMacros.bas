Attribute VB_Name = "DropDownListMacros"
Option Explicit

Sub BuildList()
Attribute BuildList.VB_ProcData.VB_Invoke_Func = " \n14"
Dim PicNum As Long
With Sheet1
    .Range("V4:W7").ClearContents 'Clear Previous Results
    On Error Resume Next
    .Shapes("SelectedIcon").Delete 'Clear Out any possible icon
    On Error GoTo 0 'Resume here on error
    .Range("P3:R15").AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("T3:T4"), CopyToRange:=.Range("V3:W3"), Unique:=True
    For PicNum = 1 To 4
            On Error Resume Next
            .Shapes("Icon" & PicNum).Delete
            On Error GoTo 0
            On Error GoTo NoPic
            With .Pictures.Insert(.Range("W" & PicNum + 3).Value)
                With .ShapeRange
                    .LockAspectRatio = msoTrue
                    .Height = 16
                    .Name = "Icon" & PicNum
                End With
            End With
                With .Shapes("Icon" & PicNum)
                    .Visible = msoFalse
                    .Left = Sheet1.Shapes("Button" & PicNum).Left
                    .Top = Sheet1.Shapes("Button" & PicNum).Top
                    .IncrementLeft 3
                    .IncrementTop 5
                End With

    Next PicNum
NoPic:
End With
End Sub
Sub ShowPopUp()
Dim PicNum As Long
With Sheet1
If .Shapes("ButtonSet").Visible = True Then 'Hide & Exit if already displayed
    HidePopUp
    Exit Sub
End If

.Shapes("ButtonSet").Visible = msoCTrue
.Shapes("OutsideBorder").Visible = msoCTrue
For PicNum = 1 To 4
On Error Resume Next
.Shapes("Icon" & PicNum).Visible = msoCTrue
On Error GoTo 0
Next PicNum
End With
End Sub
Sub HidePopUp()
Dim PicNum As Long
With Sheet1
.Shapes("ButtonSet").Visible = msoFalse
.Shapes("OutsideBorder").Visible = msoFalse
For PicNum = 1 To 4
On Error Resume Next
.Shapes("Icon" & PicNum).Visible = msoFalse
On Error GoTo 0
Next PicNum
End With
End Sub
Sub ShowSelectedIcon()
Dim IconRange As Range
Dim IconRow As Long
Dim SelOption As String
With Sheet1
    On Error Resume Next
    .Shapes("SelectedIcon").Delete
    On Error GoTo 0
    If .Range("F3").Value = Empty Then Exit Sub
    SelOption = .Range("F3").Value 'Selected option
    Set IconRange = .Range("V4:V7").Find(SelOption, , xlValues, xlWhole)
    If Not IconRange Is Nothing Then
        IconRow = IconRange.Row
        On Error GoTo NoPic
        With .Pictures.Insert(.Range("W" & IconRow).Value)
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Height = 15
            .Name = "SelectedIcon"
        End With
    End With
        With .Shapes("SelectedIcon")
            .Visible = msoCTrue
            .Left = Sheet1.Range("F3").Left
            .Top = Sheet1.Range("F3").Top
            .IncrementLeft 2
            .IncrementTop 2
        End With
    End If
NoPic:
End With
End Sub
Sub Select1()
Sheet1.Shapes("ButtonSet").ShapeStyle = msoShapeStylePreset2
Sheet1.Shapes("ButtonSet").Line.Visible = msoFalse
Sheet1.Shapes("Button1").TextFrame.Characters.Font.ColorIndex = 2
Sheet1.Shapes("Button1").ShapeStyle = msoShapeStylePreset37
Sheet1.Shapes("Button2").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button3").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button4").TextFrame.Characters.Font.ColorIndex = 1
End Sub
Sub Select2()
Sheet1.Shapes("ButtonSet").ShapeStyle = msoShapeStylePreset2
Sheet1.Shapes("ButtonSet").Line.Visible = msoFalse
Sheet1.Shapes("Button2").ShapeStyle = msoShapeStylePreset37
Sheet1.Shapes("Button2").TextFrame.Characters.Font.ColorIndex = 2
Sheet1.Shapes("Button1").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button3").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button4").TextFrame.Characters.Font.ColorIndex = 1
End Sub

Sub Select3()
Sheet1.Shapes("ButtonSet").ShapeStyle = msoShapeStylePreset2
Sheet1.Shapes("ButtonSet").Line.Visible = msoFalse
Sheet1.Shapes("Button3").ShapeStyle = msoShapeStylePreset37
Sheet1.Shapes("Button3").TextFrame.Characters.Font.ColorIndex = 2
Sheet1.Shapes("Button1").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button2").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button4").TextFrame.Characters.Font.ColorIndex = 1
End Sub

Sub Select4()
Sheet1.Shapes("ButtonSet").ShapeStyle = msoShapeStylePreset2
Sheet1.Shapes("ButtonSet").Line.Visible = msoFalse
Sheet1.Shapes("Button4").ShapeStyle = msoShapeStylePreset37
Sheet1.Shapes("Button4").TextFrame.Characters.Font.ColorIndex = 2
Sheet1.Shapes("Button1").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button2").TextFrame.Characters.Font.ColorIndex = 1
Sheet1.Shapes("Button3").TextFrame.Characters.Font.ColorIndex = 1
End Sub

