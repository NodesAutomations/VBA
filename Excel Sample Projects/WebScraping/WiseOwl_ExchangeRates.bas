Attribute VB_Name = "WiseOwl_ExchangeRates"
'@Folder("VBAProject")
Option Explicit

'Reference
'https://www.youtube.com/watch?v=dShR33CdlY8&t=2062s
Sub WiseOwl_ExchangeRate1()
    Dim IE As New InternetExplorer
    IE.Visible = True
    IE.Navigate "https://www.x-rates.com/"
    
    'Wait Until website is fully Loaded
    Do
        DoEvents
    Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    'Get HTML Doc
    Dim doc As MSHTML.HTMLDocument
    Set doc = IE.Document
    
    'Get Serchbox Input
    Dim amountElement As MSHTML.IHTMLElement
    Set amountElement = doc.getElementById("amount")
    amountElement.Value = 5
    
    Dim fromElement As MSHTML.IHTMLElement
    Set fromElement = doc.getElementById("from")
    fromElement.Value = "GBP"
    
    Dim toElement As MSHTML.IHTMLElement
    Set toElement = doc.getElementById("to")
    toElement.Value = "USD"
    
    Dim docLinks As MSHTML.IHTMLElementCollection
    Set docLinks = doc.getElementsByTagName("a")
    
    Dim docLink As MSHTML.IHTMLElement
    
    For Each docLink In docLinks
        Debug.Print docLink.getAttribute("classname"), docLink.getAttribute("href"), docLink.getAttribute("rel")
        If docLink.getAttribute("href") = "https://www.x-rates.com/table/" And docLink.getAttribute("rel") = "ratestable" Then
            docLink.Click
            Exit For
        End If
    Next
    
    IE.Quit
    Set IE = Nothing
End Sub

