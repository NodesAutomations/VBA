Attribute VB_Name = "Module1"
Option Explicit

Sub Send_Mails()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Send_Mails")
Dim i As Integer

Dim OA As Object
Dim msg As Object

Set OA = CreateObject("outlook.application")

Dim last_row As Integer
last_row = Application.CountA(sh.Range("A:A"))

For i = 2 To last_row
    Set msg = OA.createitem(0)
    msg.to = sh.Range("A" & i).Value
    msg.cc = sh.Range("B" & i).Value
    msg.Subject = sh.Range("C" & i).Value
    msg.body = sh.Range("D" & i).Value
    
    If sh.Range("E" & i).Value <> "" Then
        msg.attachments.Add sh.Range("E" & i).Value
    End If
    
    msg.send
    
    sh.Range("F" & i).Value = "Sent"

Next i

MsgBox "All the mails have been sent successfully"


End Sub
