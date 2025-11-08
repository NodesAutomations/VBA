### When to Use it
- You have a list of unique items e.g. countries, invoice numbers, customer name and addresses, project ids, product names etc.
- You need to retrieve the value of a unique item.

### Code
- Need Microsoft Scripting Runtime Reference

```vba
Sub Test()
    Dim dict As New Dictionary

    dict.Add "Vivek", 100
    dict.Add "Deven", 80
    dict.Add "Druv", 60
    
    'Count
    Debug.Print dict.Count
    
    'Acces Item or Key with Number
    Debug.Print dict.Keys(0), dict.Items(0)
    
    'Access Specific Item using Key
    Debug.Print dict("Vivek")
    
    'Check if Value Exit
    If dict.Exists("Vivek") Then
        dict("Vivek") = 90
    End If
    
    'Loop Through Dictionary
    Dim item As Variant
    For Each item In dict
        Debug.Print item, dict(item)
    Next
    
    'Remove Item
    dict.Remove ("Vivek")
    
    'Loop Using For
    Dim i As Long
    For i = 0 To dict.Count - 1
        Debug.Print dict.Keys(i), dict.Items(i)
        'use dict.Keys()(i) for compatibility with late binding
    Next i
 
End Sub
```

### Dictionary with Late Binding
```vba
Sub Test()
    Dim dict As Object
    Set dict = CreateObject("scripting.dictionary")

    dict.Add "Vivek", 100
    dict.Add "Deven", 80
    dict.Add "Druv", 60
    
    'Count
    Debug.Print dict.Count
    
    'Acces Item or Key with Number
    Debug.Print dict.Keys()(0), dict.Items()(0)
    
    'Access Specific Item using Key
    Debug.Print dict("Vivek")
    
        'Check if Value Exit
    If dict.Exists("Vivek") Then
        dict("Vivek") = 90
    End If
    
    'Access Specific Item using Key
    Debug.Print dict("Vivek")
    
    'Loop Through Dictionary
    Dim item As Variant
    For Each item In dict
        Debug.Print item, dict(item)
    Next
    
       'Remove Item
    dict.Remove ("Vivek")
    
    'Loop Using For
    Dim i As Long
    For i = 0 To dict.Count - 1
        Debug.Print dict.Keys()(i), dict.Items()(i)
    Next i
End Sub
```
### Dictionary with Array
```vb
    Dim dict   As Object
    Set dict = CreateObject("scripting.dictionary")
    dict.Add 16, Array(11, 22, 33)
    
    Debug.Print dict(16)(0)
    Debug.Print dict(16)(1)
    Debug.Print dict(16)(2)
```
