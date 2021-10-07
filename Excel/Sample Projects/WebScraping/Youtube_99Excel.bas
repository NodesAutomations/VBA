Attribute VB_Name = "Youtube_99Excel"
'@Folder("VBAProject")
Option Explicit

'Refer : https://www.youtube.com/watch?v=Erse5VJBNp0
Sub ExtractData()
    'Open IE and Navigate to Specific Page
    Dim IE As New InternetExplorer
    IE.Visible = True
    IE.Navigate ("https://www.snapdeal.com/search?keyword=iphone&santizedKeyword=&catId=&categoryId=0&suggested=false&vertical=&noOfResults=20&searchState=&clickSrc=go_header&lastKeyword=&prodCatId=&changeBackToAll=false&foundInAll=false&categoryIdSearched=&cityPageUrl=&categoryUrl=&url=&utmContent=&dealDetail=&sort=rlvncy")
    
    'Wait Until website is fully Loaded
    Do
        DoEvents
    Loop Until IE.ReadyState = READYSTATE_COMPLETE
    
    'Skip to Next if Error
    On Error Resume Next
    Dim doc As HTMLDocument
    Set doc = IE.Document
    
    Dim data As Variant
    Dim i As Integer
    For i = 0 To 99
        Set data = doc.getElementsByClassName("product-title").Item(i)
        Sheet1.Cells(i + 2, 1).Value = data.innerText
        
        Set data = doc.getElementsByClassName("lfloat product-price").Item(i)
        Sheet1.Cells(i + 2, 2).Value = data.innerText
        
        'Keep Scrolling to Load More Items
        doc.parentWindow.scrollBy 0, 99999
    Next
    IE.Quit
    Set IE = Nothing
End Sub

