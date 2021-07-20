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
