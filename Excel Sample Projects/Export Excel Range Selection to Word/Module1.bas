Attribute VB_Name = "Module1"
'@Folder("VBAProject")
Option Explicit

Sub ExcelRangeToWord()

'PURPOSE: Copy/Paste An Excel Table Into a New Word Document
'NOTE: Must have Word Object Library Active in Order to Run _
  (VBE > Tools > References > Microsoft Word 12.0 Object Library)
'SOURCE: www.TheSpreadsheetGuru.com

Dim tbl As Excel.Range
Dim WordApp As Word.Application
Dim myDoc As Word.Document
Dim WordTable As Word.Table

'Optimize Code
  Application.ScreenUpdating = False
  Application.EnableEvents = False

'Copy Range from Excel
'  Set tbl = ThisWorkbook.Worksheets(Sheet1.Name).ListObjects("Table1").Range

Set tbl = Application.Selection

'Create an Instance of MS Word
  On Error Resume Next
    
    'Is MS Word already opened?
      Set WordApp = GetObject(class:="Word.Application")
    
    'Clear the error between errors
      Err.Clear

    'If MS Word is not already open then open MS Word
      If WordApp Is Nothing Then Set WordApp = CreateObject(class:="Word.Application")
    
    'Handle if the Word Application is not found
      If Err.Number = 429 Then
        MsgBox "Microsoft Word could not be found, aborting."
        GoTo EndRoutine
      End If

  On Error GoTo 0
  
'Make MS Word Visible and Active
  WordApp.Visible = True
  WordApp.Activate
    
'Create a New Document
  Set myDoc = WordApp.Documents.Add
  
'Copy Excel Table Range
  tbl.Copy

'Paste Table into MS Word
  myDoc.Paragraphs(1).Range.PasteExcelTable _
    LinkedToExcel:=False, _
    WordFormatting:=False, _
    RTF:=False

'Autofit Table so it fits inside Word Document
 ' Set WordTable = myDoc.Tables(1)
 ' WordTable.AutoFitBehavior (wdAutoFitWindow)
   
EndRoutine:
'Optimize Code
  Application.ScreenUpdating = True
  Application.EnableEvents = True

'Clear The Clipboard
  Application.CutCopyMode = False

End Sub


