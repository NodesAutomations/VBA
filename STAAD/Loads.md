## Primary Load Cases
### Get Load Case Title
- this code with work for all loads primary loads, moving loads or load combinations
```vba
Sub Test()

    Dim objOpenSTAAD As Object
    Dim selbeamsNo As Long
    Dim SelBeams() As Long

    'Launch the OpenSTAAD Object
    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
    
    'Get Load Case Title based on Load ID
    Dim lLoadCase As Long
    Dim strLoadCaseName As String
    For lLoadCase = 1 To 3
        strLoadCaseName = objOpenSTAAD.Load.GetLoadCaseTitle(lLoadCase)
        Debug.Print strLoadCaseName
    Next lLoadCase
End Sub

```
## Load Combination
### Get Load Combination Case Count
```vba
Sub Test()

    Dim objOpenSTAAD As Object
    Dim selbeamsNo As Long
    Dim SelBeams() As Long

    'Launch the OpenSTAAD Object
    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
    
    'Load Combination Case Count
    Dim lGetLoadCombinationCaseCount As Long
    lGetLoadCombinationCaseCount = objOpenSTAAD.Load.GetLoadCombinationCaseCount
    Debug.Print lGetLoadCombinationCaseCount
End Sub
```
