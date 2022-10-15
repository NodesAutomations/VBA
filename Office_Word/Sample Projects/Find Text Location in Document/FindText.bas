Attribute VB_Name = "FindText"
'@Folder("Project")
Option Explicit

Sub CheckText()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument

    Dim MyRange As Range
    'Use Content when you want to find range in whole document
    'Set MyRange = ActiveDocument.Content
    'Use Range when you want to ignore tables,headers and footer
    Set MyRange = ActiveDocument.Range
    MyRange.Find.Execute FindText:="<<Test3dd1>>"
 
    If Not MyRange.Find.Found Then
        MsgBox MyRange.Find.Text
    End If
    'MyRange.Paste
   
End Sub

Sub FindTextPosition()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument

    Dim MyRange As Range
    Set MyRange = ActiveDocument.Range
    With ActiveDocument.Range.Find
        .Text = "<<Test1>>"
        .Execute
    
    End With
    ' MyRange.Find.Execute FindText:="<<Test11>>"
 
    'MsgBox "Position = " & MyRange.Text
 
End Sub

Sub findTextInsideTable()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument
    Dim tbl As Table
    Dim MyRange As Range
    For Each tbl In wdDoc.Tables
        Set MyRange = tbl.Range
    
        MyRange.Find.Execute FindText:="<<Test6>>"
        
        If MyRange.Start > 0 Then
            MsgBox "Position = " & MyRange.Start
            'MyRange.Paste
        End If
    Next
    
End Sub

Sub FindTableTags()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument
    Dim tbl As Table
    Dim MyRange As Range
    
    Dim i As Integer, j As Integer
    
    For i = 1 To wdDoc.Tables.Count
        Set tbl = wdDoc.Tables(i)
        If tbl.Columns.Count = 1 Then
            'Debug.Print tbl.Rows.Count, tbl.Columns.Count
            For j = 1 To tbl.Rows.Count
                Set MyRange = tbl.Cell(j, 1).Range
                If InStr(MyRange.Text, "<<Test2>>") > 0 Then
                    tbl.Cell(j, 1).Range.Text = "Match Found"
                End If
            Next
        End If
    Next
    
End Sub

