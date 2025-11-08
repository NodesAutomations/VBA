Attribute VB_Name = "Module1"
Sub Bestsellers()
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
Set digisht = ThisWorkbook.Sheets("Bestsellers - Digital")
Set physsht = ThisWorkbook.Sheets("Bestsellers - Physical")

On Error Resume Next

digisht.Range("B2", "I101").ClearContents
physsht.Range("B2", "L101").ClearContents



Dim html As HTMLDocument
Set IE = New InternetExplorer


str = "https://www.amazon.co.uk/Best-Sellers-Kindle-Store/zgbs/digital-text/"


IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document
counter = 2
For Each element In Doc.getElementsByClassName("zg-item-immersion")
    digisht.Cells(counter, 2) = CStr(Mid(element.innerHTML, InStr(1, element.innerHTML, "/dp/") + 4, 10))
    counter = counter + 1
Next

str = "https://www.amazon.co.uk/Best-Sellers-Kindle-Store/zgbs/digital-text/?pg=2"


IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document
counter = 52
For Each element In Doc.getElementsByClassName("zg-item-immersion")
    digisht.Cells(counter, 2) = Mid(element.innerHTML, InStr(1, element.innerHTML, "/dp/") + 4, 10)
    counter = counter + 1
Next

str = "https://www.amazon.co.uk/Best-Sellers-Books/zgbs/books"


IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document
counter = 2
For Each element In Doc.getElementsByClassName("zg-item-immersion")
    physsht.Cells(counter, 2) = CStr(Mid(element.innerHTML, InStr(1, element.innerHTML, "/dp/") + 4, 10))
    counter = counter + 1
Next

str = "https://www.amazon.co.uk/Best-Sellers-Books/zgbs/books?pg=2"


IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document
counter = 52
For Each element In Doc.getElementsByClassName("zg-item-immersion")
    physsht.Cells(counter, 2) = CStr(Mid(element.innerHTML, InStr(1, element.innerHTML, "/dp/") + 4, 10))
    counter = counter + 1
Next

For i = 2 To 101

ASIN1 = digisht.Cells(i, 2)

str = "https://www.amazon.co.uk/dp/" & ASIN1

IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document



If InStr(1, Doc.getElementById("CombinedBuybox").innerHTML, "kindle-unlimited") > 0 Then
    digisht.Cells(i, 11) = True
End If


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
        digisht.Cells(i, 7) = Application.NumberValue(Trim(Right(elementa.innerText, Len(elementa.innerText) - 1)))
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


    
For i = 2 To 101

ASIN1 = physsht.Cells(i, 2)

str = "https://www.amazon.co.uk/dp/" & ASIN1

IE.navigate str
IE.Visible = True
 
Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document



physsht.Cells(i, 4) = Trim(Doc.getElementById("productTitle").innerText)

physsht.Cells(i, 8) = Application.NumberValue(Trim(Mid(Doc.getElementById("buyNewSection").innerText, InStr(1, Doc.getElementById("buyNewSection").innerText, "£") + 1, Len(Doc.getElementById("buyNewSection").innerText) - InStr(1, Doc.getElementById("buyNewSection").innerText, "£"))))

If physsht.Cells(i, 8) = "" Then
    For Each element In Doc.getElementsByClassName("a-color-price")
        physsht.Cells(i, 8) = Application.NumberValue(Trim(Mid(element.innerText, InStr(1, element.innerText, "£") + 1, Len(element.innerText) - InStr(1, element.innerText, "£"))))
    Next
End If
        


counter = 0
For Each element In Doc.getElementsByClassName("author notFaded")
    For Each elementa In element.getElementsByClassName("a-declarative")
         If counter = 0 Then
            physsht.Cells(i, 5) = elementa.innerText
            counter = 1
        End If
    Next
Next

For Each element In Doc.getElementById("buyBoxInner").getElementsByTagName("li")
    If InStr(1, element.innerText, "RP:") > 0 Then
        physsht.Cells(i, 9) = Application.NumberValue(Trim(Mid(element.innerText, InStr(1, element.innerText, "£") + 1, Len(element.innerText) - InStr(1, element.innerText, "£"))))
    End If
Next

If Len(Doc.getElementById("availability").innerText) > 100 Then
    physsht.Cells(i, 12) = "out of stock"
Else
    physsht.Cells(i, 12) = Trim(Doc.getElementById("availability").innerText)
End If

counter = 0
For Each element In Doc.getElementsByClassName("content")
    For Each elementa In element.getElementsByClassName("a-icon-alt")
        physsht.Cells(i, 10) = Left(elementa.innerText, 3)
    Next
    For Each elementa In element.getElementsByClassName("a-link-normal")
        If Len(elementa.innerText) > 0 Then
            physsht.Cells(i, 11) = Left(Trim(elementa.innerText), InStr(1, Trim(elementa.innerText), " ") - 1)
        End If
    Next
    For Each elementa In element.getElementsByTagName("li")
        If InStr(1, elementa.innerText, "ublish") > 0 Then
            If Len(elementa.innerText) - Len(Replace(elementa.innerText, "(", "")) > 1 Then
                ASIN2 = Right(elementa.innerText, Len(elementa.innerText) - InStr(1, elementa.innerText, "(") - 1)
                physsht.Cells(i, 7) = Replace(Trim(Mid(ASIN2, InStr(1, ASIN2, "(") + 1, Len(ASIN2) - InStr(1, ASIN2, "(") - 1)), ".", "")
                Else
                physsht.Cells(i, 7) = Replace(Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, "(") + 1, Len(elementa.innerText) - InStr(1, elementa.innerText, "(") - 1)), ".", "")
            End If
            physsht.Cells(i, 6) = Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, ":") + 1, InStr(1, elementa.innerText, "(") - InStr(1, elementa.innerText, ":") - 1))
        End If
        If InStr(1, elementa.innerText, "old b") > 0 Then
            physsht.Cells(i, 6) = Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, ":") + 1, Len(elementa.innerText) - InStr(1, elementa.innerText, ":") - 1))
        End If
        If InStr(1, elementa.innerText, "SBN-13") > 0 Then
            physsht.Cells(i, 3) = Replace(Trim(Mid(elementa.innerText, InStr(1, elementa.innerText, ":") + 1, Len(elementa.innerText) - InStr(1, elementa.innerText, ":"))), "-", "")
        End If
    Next
Next
If InStr(1, physsht.Cells(i, 11), "e") > 0 Then
    physsht.Cells(i, 11) = 0
    physsht.Cells(i, 10) = "N/A"
End If

Next
    




End Sub


