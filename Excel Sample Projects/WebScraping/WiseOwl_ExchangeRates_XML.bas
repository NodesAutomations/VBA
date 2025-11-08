Attribute VB_Name = "WiseOwl_ExchangeRates_XML"
'@Folder("VBAProject")
Option Explicit

'Reference
'https://www.youtube.com/watch?v=dShR33CdlY8&t=2062s
Sub WiseOwl_ExchangeRate_String()
    Dim xmlDoc As New MSXML2.XMLHTTP60
    xmlDoc.Open "GET", "https://www.x-rates.com/table/?from=GBP&amount=5", False
    xmlDoc.send
    
 
    'Get HTML Doc
    Dim doc As New MSHTML.HTMLDocument
    doc.body.innerHTML = xmlDoc.responseText
    
    Call processHTML(doc)
    
End Sub

Sub processHTML(page As MSHTML.HTMLDocument)
    Dim htmlTables As MSHTML.IHTMLElementCollection
    Dim htmlTable As MSHTML.IHTMLElement
    Dim htmlRow As MSHTML.IHTMLElement
    Dim htmlCell As MSHTML.IHTMLElement
    Dim rowNum As Long, colNum As Integer
    
    Set htmlTables = page.getElementsByTagName("table")
    
    For Each htmlTable In htmlTables
        Worksheets.Add
        Range("A1").Value = htmlTable.className
        Range("B1").Value = Now
        
        rowNum = 2
        For Each htmlRow In htmlTable.getElementsByTagName("tr")
            'Debug.Print vbTab & htmlRow.innerText
            colNum = 1
            For Each htmlCell In htmlRow.Children
                Cells(rowNum, colNum) = htmlCell.innerText
                colNum = colNum + 1
            Next
            rowNum = rowNum + 1
        Next
    Next
End Sub

