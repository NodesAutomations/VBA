Attribute VB_Name = "ExtractData"
'@Folder("VBAProject")
Option Explicit

Sub GetData()
    Dim pdfPath As String
    pdfPath = "C:\Users\Ryzen2600x\source\repos\Personal_Vivek_PDF_Scraper\data.pdf"

     
    'Start Word Application
    Dim wdApp As New Word.Application
    wdApp.Visible = True
    
    'Create Word Document from template
    Dim wdDoc As Word.Document
    Set wdDoc = wdApp.Documents.Open(pdfPath, False, True, Format:="PDF Files")
    
    Dim invoiceText As String
    invoiceText = wdDoc.Tables(1).Cell(1, 3).Range.Text
    
    'Invoice Date
    Sheet1.Range("B1") = Left(invoiceText, 10)
     
    'Invoice Number
    Sheet1.Range("B2") = Mid(invoiceText, 11, 16)
     
    Dim totalText As String
    totalText = wdDoc.Tables(2).Cell(4, 2).Range.Text
    'Total
    Sheet1.Range("B3") = Left(totalText, Len(totalText) - 3)
    
    'Item
    Sheet1.Range("B4") = wdDoc.Tables(2).Cell(2, 2).Range.Text
    
    'Close Word Document
    wdDoc.Close
    'Close Word App
    wdApp.Quit
    'Release
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub
End Sub

