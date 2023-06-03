### Sample code to get fii and dii data using webscrapping
```vba
Option Explicit

Sub ExtractTables()
    'declare variables
    Dim ie As Object
    Dim doc As Object
    Dim tables As Object
    Dim table As Object
    Dim rows As Object
    Dim row As Object
    Dim cells As Object
    Dim cell As Object
    Dim r As Long
    Dim c As Long
    
    'create a new Internet Explorer instance
    Set ie = CreateObject("InternetExplorer.Application")
    
    'navigate to the website
    ie.Navigate "https://www.nseindia.com/reports/fii-dii"
    
    'wait until the page is fully loaded
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop
    
    'get the document object
    Set doc = ie.Document
    
    'get all the tables in the document
    Set tables = doc.getElementsByTagName("table")
    
    'loop through each table
    For Each table In tables
        
        'get the rows in the table
        Set rows = table.getElementsByTagName("tr")
        
        'initialize the row counter
        r = 1
        
        'loop through each row
        For Each row In rows
            
            'get the cells in the row
            Set cells = row.getElementsByTagName("td")
            
            'initialize the column counter
            c = 1
            
            'loop through each cell
            For Each cell In cells
                
                'write the cell value to the worksheet
                ActiveSheet.cells(r, c).Value = cell.innerText
                
                'increment the column counter
                c = c + 1
                
            Next cell
            
            'increment the row counter
            r = r + 1
            
        Next row
        
        'move to the next empty column for the next table
        Do While ActiveSheet.cells(1, c) <> ""
            c = c + 1
        Loop
        
    Next table
    
    'close the Internet Explorer instance
    ie.Quit
    
    'release the variables
    Set ie = Nothing
    Set doc = Nothing
    Set tables = Nothing
    Set table = Nothing
    Set rows = Nothing
    Set row = Nothing
    Set cells = Nothing
    Set cell = Nothing
    
End Sub
```
