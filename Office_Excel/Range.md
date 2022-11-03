### Find Last USed Row Column
```vba
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.range("A" & ws.Rows.Count).End(xlUp).Row
```
### Find Current Region based on specific cell
```vba
Sheet1.Range("E8").CurrentRegion.Address
'Print
'$D$8:$E$10
```
![image](https://user-images.githubusercontent.com/60865708/126315441-33bb5d22-5478-4337-b892-d6561fad3103.png)

### Get Range Data in Arra
```vba
Dim arr as Variant
arr=Sheet1.Range("E8").CurrentRegion.Value
```
### Store Range In Arrray For Faster Processing
```vba
dim rg as Range
set rg=Sheet1.Range("E8").CurrentRegion

'Reading Data From Range
dim arr as Variant
arr=rg.Value

'Writing Data from Range
Sheet1.Range("I8:M18").Value=arr
```
### Name Range
```visual-basic
Sub NameRange_Add()

Dim cell As Range
Dim rng As Range
Dim RangeName As String
Dim CellName As String

'Single Cell Reference (Workbook Scope)
  RangeName = "Price"
  CellName = "D7"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell

'Single Cell Reference (Worksheet Scope)
  RangeName = "Year"
  CellName = "A2"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  Worksheets("Sheet1").Names.Add Name:=RangeName, RefersTo:=cell

'Range of Cells Reference (Workbook Scope)
  RangeName = "myData"
  CellName = "F9:J18"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell

'Secret Named Range (doesn't show up in Name Manager)
  RangeName = "Username"
  CellName = "L45"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell, Visible:=False

End Sub
```

```visual-basic
Sub Update()

Dim cell As Range
Dim CellName As String
Dim RangeName As String
Dim RangeValue As String
Dim SheetName As String

Dim FileNum As Integer
Dim DataLine As String

FileNum = FreeFile()
Open "C:\Users\Ryzen2600x\TestRepo\AddNamedRange\AddNamedRange\MissingVar.txt" For Input As #FileNum
Dim str() As String

While Not EOF(FileNum)
    Line Input #FileNum, DataLine ' read in data 1 line at a time
    str = Split(DataLine, ",")
    'MsgBox (str(2))
        
    SheetName = str(0)
    CellName = str(1)
    RangeName = str(2)
    RangeValue = str(3)

 Set cell = Worksheets(SheetName).Range(CellName)
 If Not RangeName = "" Then
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell
  Range(RangeName) = RangeValue
  Else
  cell.Value = RangeValue
 End If
Wend
```
