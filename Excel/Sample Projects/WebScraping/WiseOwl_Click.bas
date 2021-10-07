Attribute VB_Name = "WiseOwl_Click"
'@Folder("VBAProject")
Option Explicit

'Reference
'https://www.youtube.com/watch?v=dShR33CdlY8&t=2062s
Sub WiseOwl()
    Dim IE As New InternetExplorer
    IE.Visible = True
    IE.Navigate "https://en.wikipedia.org/wiki/Main_Page"
    
    'Wait Until website is fully Loaded
    Do
        DoEvents
    Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    Debug.Print IE.LocationName, IE.LocationURL
    
    'Update Searchbox
    IE.Document.forms("searchform").elements("search").Value = "Document Object Model"
    IE.Document.forms("searchform").elements("go").Click
    
    IE.Quit
    Set IE = Nothing
End Sub


Sub WiseOwl_GetData()
    Dim IE As New InternetExplorer
    IE.Visible = True
    IE.Navigate "wiseowl.co.uk"
    
    'Wait Until website is fully Loaded
    Do
        DoEvents
    Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    Dim doc As MSHTML.HTMLDocument
    Set doc = IE.Document
    
    'Get Serchbox Input
    Dim element As MSHTML.IHTMLElement
    
    'By Id
    Set element = doc.getElementById("what")
    element.Value = "Excel VBA"
    
    Dim elements As MSHTML.IHTMLElementCollection
    Set elements = doc.getElementsByTagName("button")
    
    Dim tempelement As MSHTML.IHTMLElement
    
    For Each tempelement In elements
        Debug.Print tempelement.className, tempelement.tagName, tempelement.ID, tempelement.innerText
        
    Next
    elements(0).Click
    
    'doc.get
    IE.Quit
    Set IE = Nothing
End Sub


