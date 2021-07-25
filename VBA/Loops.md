## For Next

### Code Snippet
```vba
  For Counter = Start To End [Step Value]
  [Code Block to Execute]
  Next [counter]
```

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

##
