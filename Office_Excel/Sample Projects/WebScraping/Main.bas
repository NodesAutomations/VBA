Attribute VB_Name = "Main"
'@Folder("VBAProject")
Option Explicit

Sub Test()
    'Create Instance for Internet Explorer
    Dim ie As New InternetExplorer
    ie.Visible = True
    ie.Navigate "http://demo.guru99.com/test/web-table-element.php"

    Do
        DoEvents
    Loop Until ie.ReadyState = READYSTATE_COMPLETE

    'Store Data to Doc
    Dim doc As New HTMLDocument
    Set doc = ie.Document

    Dim ecoll As Object
    Set ecoll = doc.getElementsByTagName("table")

    
End Sub

