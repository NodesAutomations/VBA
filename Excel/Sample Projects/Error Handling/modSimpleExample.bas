Attribute VB_Name = "modSimpleExample"
Option Explicit

'

' Level 1
' https://excelmacromastery.com/
' Place the cursor in this sub and press F5 to see error
Public Sub TopMost()

    On Error GoTo eh
    
    Call Level2

done:
    Exit Sub
eh:
    DisplayError Err.Source, Err.Description, "modSimpleExample.TopMost", Erl
End Sub

' Level 2
' https://excelmacromastery.com/
Private Sub Level2()

    On Error GoTo eh
    
    Call Level3
     
done:
    Exit Sub
eh:
    RaiseError Err.Number, Err.Source, "modSimpleExample.Level2", Err.Description, Erl
End Sub

' Level 2
' https://excelmacromastery.com/
Private Sub Level3()

    On Error GoTo eh
    
    ' TYPE MISMATCH ERROR
    Dim total As Long
    total = "a"
    

done:
    Exit Sub
eh:
    RaiseError Err.Number, Err.Source, "modSimpleExample.Level3", Err.Description, Erl
End Sub


