### Store Custom Settings at document level
```vba
Sub test()
Dim test As String
test = ThisWorkbook.CustomDocumentProperties("TestMe")
Debug.Print test
ThisWorkbook.CustomDocumentProperties("TestMe") = "Updated Path"
Debug.Print test
End Sub
```
