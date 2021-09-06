## If 

### Code Syntax
```vba
If <Condition> Then <Statement>
```
```vba
If <Condition> Then
   Statements1
Else
   Statements2
End if
```
```vba
If <Condition> Then
   Statements1
Else
   Statements2
End if
```
## IIF 
### Code syntax
```vba
Function CheckIt (TestMe As Integer)
    CheckIt = IIf(TestMe > 1000, "Large", "Small")
End Function
```
### Select Case

### Code Syntax
```vba
Select Case Expression
    Case Expression1
	Statement1
    Case Expression2
        Statement2
    Case ExpressionN
        StatementN
End Select 
```
Multiple Cases
```vba
' https://excelmacromastery.com/
Public Sub Select_Case_Multi()

    Dim city As String
    ' Change value to test
    city = "Dublin"
    
    ' Print the name of the airport based on the code
    Select Case city
        Case "Paris", "London", "Dublin"
            Debug.Print "Europe"
        Case "Singapore", "Hanoi"
            Debug.Print "Asia"
        Case Else
            MsgBox "The city is not valid.", vbInformation
    End Select

End Sub
```
