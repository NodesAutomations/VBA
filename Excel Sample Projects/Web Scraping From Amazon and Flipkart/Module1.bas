Attribute VB_Name = "Module1"
Option Explicit

Sub Search()


Dim sh As Worksheet
Set sh = ActiveSheet

sh.Range("A1").Value = 1

Call Fetch_from_Amazon
Call Fetch_from_Flipkart

MsgBox "Done"

End Sub

Sub Fetch_from_Amazon()
 
Dim sh As Worksheet
Set sh = ActiveSheet


Dim i As Integer

Dim IE As InternetExplorer
Dim html_doc As HTMLDocument

Set IE = New InternetExplorer

IE.Visible = True
IE.navigate "WWW.Amazon.in"


Do Until IE.readyState = READYSTATE_COMPLETE
    DoEvents
Loop



Set html_doc = IE.document

For i = 4 To sh.Range("A" & Application.Rows.Count).End(xlUp).Row
    DoEvents
    
    If sh.Range("A1").Value = 0 Then Exit Sub
    
    On Error Resume Next
    
    html_doc.getElementById("twotabsearchtextbox").Value = sh.Range("A" & i).Value
    html_doc.getElementsByClassName("nav-input")(1).Click
    
    Do Until IE.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Application.Wait Now + TimeValue("0:00:02")
    
    sh.Range("B" & i).Value = html_doc.getElementsByClassName("a-size-medium a-color-base a-text-normal")(0).innerText
    sh.Range("C" & i).Value = html_doc.getElementsByClassName("a-price-whole")(0).innerText
     
Next i

IE.Quit


End Sub

Sub Fetch_from_Flipkart()

Dim sh As Worksheet
Set sh = ActiveSheet

Dim IE As InternetExplorer
Dim html_doc As HTMLDocument
Dim i As Integer

Set IE = New InternetExplorer

IE.Visible = True
IE.navigate "WWW.Flipkart.com"


Do Until IE.readyState = READYSTATE_COMPLETE
    DoEvents
Loop

Set html_doc = IE.document

For i = 4 To sh.Range("A" & Application.Rows.Count).End(xlUp).Row

    If sh.Range("A1").Value = 0 Then Exit Sub
    
    DoEvents
    On Error Resume Next

    html_doc.getElementsByClassName("_3704LK")(0).Value = sh.Range("A" & i).Value
    html_doc.getElementsByClassName("L0Z3Pu")(0).Click
    
    Do Until IE.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Application.Wait Now + TimeValue("0:00:03")
    
    sh.Range("D" & i).Value = html_doc.getElementsByClassName("_4rR01T")(0).innerHTML
    sh.Range("E" & i).Value = html_doc.getElementsByClassName("_30jeq3 _1_WHN1")(0).innerHTML

Next i

IE.Quit

End Sub


Sub Clear_Sheet()

Dim sh As Worksheet
Set sh = ActiveSheet

sh.Range("B4:E" & Application.Rows.Count).ClearContents

End Sub


Sub Stop_Macro()

Dim sh As Worksheet
Set sh = ActiveSheet

sh.Range("A1").Value = 0

End Sub
