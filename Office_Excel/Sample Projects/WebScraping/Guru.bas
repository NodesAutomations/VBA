Attribute VB_Name = "Guru"
'@Folder("VBAProject")
Option Explicit

Sub Test()
    'Create Instance for Internet Explorer
    Dim IE As New InternetExplorer
    IE.Visible = True
    IE.Navigate "http://demo.guru99.com/test/web-table-element.php"

    Do
        DoEvents
    Loop Until IE.ReadyState = READYSTATE_COMPLETE
codw
    'Store Data to Doc
    Dim doc As New HTMLDocument
    Set doc = IE.Document

    Dim ecoll As Object
    Set ecoll = doc.getElementsByTagName("table")

    
End Sub

