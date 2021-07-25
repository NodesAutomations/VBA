## For Next

### Code Syntax
```vba
  For Counter = Start To End [Step Value]
  [Code Block to Execute]
  Next [counter]
```
### Examples
Basic Use
```vba
Sub AddNumbers()
    Dim Total As Integer
    Dim Count As Integer
    Total = 0
    For Count = 1 To 10
        Total = Total + Count
    Next Count
    MsgBox Total
End Sub
```
With Step
- you can Jump Any Number of Loops using Steps
- Use Negetive for Reverse Loops
```vba
Sub AddEvenNumbers()
    Dim Total As Integer
    Dim Count As Integer
    Total = 0
    For Count = 2 To 10 Step 2
        Total = Total + Count
    Next Count
    MsgBox Total
End Sub
```
Exit For
```vba
Sub HghlightNegative()
    Dim Rng As range
    Set Rng = range("A1", range("A1").End(xlDown))
    Counter = Rng.Count
    For i = 1 To Counter
        If WorksheetFunction.Min(Rng) >= 0 Then
            Exit For
        End If
        If Rng(i).Value < 0 Then
            Rng(i).Font.Color = vbRed
        End If
    Next i
End Sub
```

## Do While

### Code Syntax
```vba
Do [While condition]
[Code block to Execute]
Loop
```
```vba
Do
[Code block to Execute]
Loop [While condition]
```
### Examples
General Use
```vba
Sub AddFirst10PositiveIntegers()
    Dim i As Integer
    i = 1
    Do While i <= 10
        Result = Result + i
        i = i + 1
    Loop
    MsgBox Result
End Sub
```
Exit Do
```vba
Sub EnterCurrentMonthDates()
    Dim CMDate As Date
    Dim i As Integer
    i = 0
    CMDate = DateSerial(Year(Date), Month(Date), 1)
    Do While Month(CMDate) = Month(Date)
        range("A1").Offset(i, 0) = CMDate
        i = i + 1
        If i >= 10 Then
            Exit Do
        End If
        CMDate = CMDate + 1
    Loop
End Sub
```
## Do Until

### Code Syntax
```vba
Do [Until condition]
[Code block to Execute]
Loop
```
```vba
Do
[Code block to Execute]
Loop [Until condition]
```
Behave Same us Do while

## For Each

### Basic Syntax
```vba
For Each element In collection
[Code Block to Execute]
Next [element]
```
### Examples
Basic Usage
```vba
Sub SaveAllWorkbooks()
    Dim wb As Workbook
    For Each wb In Workbooks
        wb.Save
    Next wb
End Sub
```
- Use the ‘Exit For’ statement in the For Each-Next loop to come out of the loop. 


