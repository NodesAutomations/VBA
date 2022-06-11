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
### Get Primary Count
```vba
Sub Test()

    Dim objOpenSTAAD As Object
    Dim selbeamsNo As Long
    Dim SelBeams() As Long

    'Launch the OpenSTAAD Object
    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
    
    'Get Primary Load Count
    Dim lPrimaryLoadCaseCount  As Long
    lPrimaryLoadCaseCount = objOpenSTAAD.Load.GetPrimaryLoadCaseCount
    Debug.Print lPrimaryLoadCaseCount
    
End Sub
```
### Get All LoadID with Moving Loads
```vba
'@Folder("VBAProject")
Option Explicit

Sub Test()
    Dim objOpenSTAAD As Object
    Dim selbeamsNo As Long
    Dim SelBeams() As Long

    'Launch the OpenSTAAD Object
    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
    
    'Get Primary Load Count
    Dim lPrimaryLoadCaseCount  As Long
    lPrimaryLoadCaseCount = objOpenSTAAD.Load.GetPrimaryLoadCaseCount
    'Debug.Print lPrimaryLoadCaseCount
    
    'Get Primary Load Case Array
    Dim lPrimaryLoadCaseNumbersArray() As Long
    ReDim lPrimaryLoadCaseNumbersArray(lPrimaryLoadCaseCount - 1)
    'Get Primary Load Case Numbers
    objOpenSTAAD.Load.GetPrimaryLoadCaseNumbers lPrimaryLoadCaseNumbersArray
 
    Dim strLoadCaseName As String
    
    Dim i As Integer
    Dim lastLoadID As Long
    For i = 0 To UBound(lPrimaryLoadCaseNumbersArray)
        If lPrimaryLoadCaseNumbersArray(i) = 0 Then
            lPrimaryLoadCaseNumbersArray(i) = GetNextLoadID(lastLoadID, objOpenSTAAD)
        End If
        strLoadCaseName = objOpenSTAAD.Load.GetLoadCaseTitle(lPrimaryLoadCaseNumbersArray(i))
        Debug.Print strLoadCaseName
        lastLoadID = lPrimaryLoadCaseNumbersArray(i)
    Next
End Sub

Private Function GetNextLoadID(lastLoadID As Long, objOpenSTAAD As Object) As Long
    Dim strLoadCaseName As String
    Dim loadID As Long
    loadID = lastLoadID + 1
    On Error Resume Next
    Do While GetNextLoadID = 0
        strLoadCaseName = objOpenSTAAD.Load.GetLoadCaseTitle(loadID)
        If strLoadCaseName <> "" Then
            If InStr(strLoadCaseName, "#") > 0 Then
                GetNextLoadID = CLng(Mid(strLoadCaseName, InStr(strLoadCaseName, "#") + 1, InStr(strLoadCaseName, ", (") - InStr(strLoadCaseName, "#") - 1))
                Exit Function
            End If
        End If
        loadID = loadID + 1
    Loop
End Function
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
