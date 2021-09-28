Attribute VB_Name = "Module2"
Sub digiscrape()
Dim IE As New InternetExplorer
Dim str As String
Dim Doc As HTMLDocument
Dim tagElements As Object
Dim element As HTMLObjectElement
Dim elementa As HTMLObjectElement
Dim elementb As HTMLObjectElement
Dim ASIN1 As String
Dim ASIN2 As String
Dim LastRow As Long
Dim counter As Integer
Dim asincount As Long
Dim titlecount As Long
Set digisht = ThisWorkbook.Sheets("Manual Scrape - Digital")


On Error Resume Next

digisht.Range("A2", "A20001").Clear
digisht.Range("C2", "I20001").Clear

Dim html As HTMLDocument
Set IE = New InternetExplorer

LastRow = digisht.Range("B2").End(xlDown).Row
MsgBox LastRow

For i = 2 To LastRow

ASIN1 = digisht.Cells(i, 2)

str = "https://www.amazon.co.uk/dp/" & ASIN1

IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document



digisht.Cells(i, 3) = Trim(Doc.getElementById("ebooksProductTitle").innerText)

counter = 0
For Each element In Doc.getElementsByClassName("author notFaded")
    For Each elementa In element.getElementsByClassName("a-declarative")
         If counter = 0 Then
            digisht.Cells(i, 4) = elementa.innerText
            counter = 1
        End If
    Next
Next

For Each element In Doc.getElementsByClassName("kindle-price")
    For Each elementa In element.getElementsByClassName("a-size-medium a-color-price")
        digisht.Cells(i, 7) = Application.NumberValue(Right(elementa.innerText, Len(elementa.innerText) - 1))
    Next
Next
counter = 0
For Each element In Doc.getElementsByClassName("content")
    For Each elementa In element.getElementsByClassName("a-icon-alt")
        digisht.Cells(i, 8) = Left(elementa.innerText, 3)
    Next
    For Each elementa In element.getElementsByClassName("a-link-normal")
        If Len(elementa.innerText) > 0 Then
            digisht.Cells(i, 9) = Left(Trim(elementa.innerText), InStr(1, Trim(elementa.innerText), " ") - 1)
        End If
    Next
    For Each elementa In element.getElementsByTagName("li")
        If InStr(1, elementa.innerText, "Rank") > 0 Then
            digisht.Cells(i, 1) = Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, "#") + 1, InStr(InStr(1, elementa.innerText, "#"), elementa.innerText, " ") - InStr(1, elementa.innerText, "#")))
        End If
    Next
    For Each elementa In element.getElementsByTagName("li")
        If InStr(1, elementa.innerText, "old b") > 0 Then
            digisht.Cells(i, 5) = Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, ":") + 1, Len(elementa.innerText) - InStr(1, elementa.innerText, ":") - 1))
        End If
        If InStr(1, elementa.innerText, "ublish") > 0 Then
            digisht.Cells(i, 6) = Replace(Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, "(") + 1, Len(elementa.innerText) - InStr(1, elementa.innerText, "(") - 1)), ".", "")
            digisht.Cells(i, 5) = Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, ":") + 1, InStr(1, elementa.innerText, "(") - InStr(1, elementa.innerText, ":") - 1))
            GoTo escape
        End If
    Next
Next
escape:
If InStr(1, digisht.Cells(i, 9), "e") > 0 Then
    digisht.Cells(i, 9) = 0
    digisht.Cells(i, 8) = "N/A"
End If


Next
End Sub

Sub physcrape()

Dim IE As New InternetExplorer
Dim str As String
Dim Doc As HTMLDocument
Dim tagElements As Object
Dim element As HTMLObjectElement
Dim elementa As HTMLObjectElement
Dim elementb As HTMLObjectElement
Dim ASIN1 As String
Dim ASIN2 As String
Dim LastRow As Long
Dim counter As Integer
Dim asincount As Long
Dim titlecount As Long
Set insht = ThisWorkbook.Sheets("Input")
Set outsht = ThisWorkbook.Sheets("Output")


On Error Resume Next

Dim html As HTMLDocument
Set IE = New InternetExplorer

LastRow = insht.Range("A2").End(xlDown).Row
MsgBox LastRow

For i = 2 To LastRow

ASIN1 = insht.Cells(i, 1)

outsht.Cells(i, 1) = ASIN1

str = "https://play.google.com/store/search?q=" & ASIN1

IE.navigate str
IE.Visible = True
 
Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document


For Each element In Doc.getElementsByClassName("LCATme")
    counter = 0
    For Each elementa In element.getElementsByTagName("span")
        counter = counter + 1
    Next
    If counter = 1 Then
        outsht.Cells(i, 3) = Replace(element.innerText, "£", "")
    Else
        counter = 2
        countera = 0
        For Each elementa In element.getElementsByTagName("span")
            outsht.Cells(i, counter) = Replace(elementa.innerText, "£", "")
            If countera = 1 Then
                counter = counter + 1
                countera = 0
            Else
            countera = countera + 1
            End If
        Next
    End If
Next

Next



End Sub
