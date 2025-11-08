## Use Advance Filter without Copy any data

```vba
Sub AdvanceFilterInPlace()
    
    Dim dataRange As Range
    Dim criteriaRange As Range

    Set dataRange = Sheet1.Range("A4").CurrentRegion
    Set criteriaRange = Sheet1.Range("A1").CurrentRegion
    
    dataRange.AdvancedFilter xlFilterInPlace, criteriaRange
End Sub
```

### Use Advance Filter with Copy Data to another range
```vba
Sub AdvanceFilterCopy()
    'Clear Previous Data First
    Sheet1.Range("G1").CurrentRegion.ClearContents
    
    Dim dataRange As Range
    Dim criteriaRange As Range
    Dim outputRange As Range
    
    Set dataRange = Sheet1.Range("A4").CurrentRegion
    Set criteriaRange = Sheet1.Range("A1").CurrentRegion
    
    'We Can Also Copy data to another sheet
    Set outputRange = Sheet1.Range("G1")
    dataRange.AdvancedFilter xlFilterCopy, criteriaRange, outputRange
End Sub
```

References 
- [Macro Mastery](https://www.youtube.com/watch?v=0YNhxVu2a5s)
