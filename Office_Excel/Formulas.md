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
