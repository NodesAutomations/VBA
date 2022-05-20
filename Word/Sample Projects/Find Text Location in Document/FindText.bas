Attribute VB_Name = "FindText"
'@Folder("Project")
Option Explicit

Sub FindText()
    Dim wdDoc As Document
    Set wdDoc = ActiveDocument

    Dim MyRange As Range
    Set MyRange = ActiveDocument.Range
    MyRange.Find.Execute FindText:="<<Test1>>"
 
    MsgBox "Position = " & MyRange.Text
    MyRange.Paste
   
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

