# Remove PageNumber From Specific Section Header and Footer

```vba
Sub DeletePageNumbers()

Dim objSect As Section
Dim objHF As HeaderFooter
Dim objPNum As PageNumber

For Each objSect In ActiveDocument.Sections
    
    For Each objHF In objSect.Headers
        For Each objPNum In objHF.PageNumbers
            objPNum.Delete
        Next
    Next
    
    For Each objHF In objSect.Footers
        For Each objPNum In objHF.PageNumbers
            objPNum.Delete
        Next
    Next
    
Next

End Sub
```
