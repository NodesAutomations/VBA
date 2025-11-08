### Formula for DropDown Datavalidation.
For Source Table with Single Column or First Column as Data Validation
```
=INDIRECT("BrandTable")
```
For Source Table with Multiple Column Or Intermediate Colun as Data Validation
```
=INDIRECT("AudioCategoryTable[Category]")
```
### VBA
Apply Same Formula to range
```vba
Range("L5").Formula = "=IF('Utensils-Portions'!A2="""","""",'Utensils-Portions'!A2)"
Range("L5").AutoFill Destination:=Range("L5:L106")
```

### formul to find Values from table using Vlookup
```vba
  'Label
            .Label = dataTable.Cells(i, sheet.Range("D1").Column)
            
            Dim labelTable As ListObject
            Set labelTable = SettingsSheet.ListObjects("LabelTable")
            
            Dim dataRange As Range
            Set dataRange = labelTable.Range.Columns(1)
            If WorksheetFunction.CountIf(dataRange, .Label) = 1 Then
                .Label = WorksheetFunction.VLookup(.Label, SettingsSheet.Range("LabelTable[#All]"), 2, False)
            End If
```
