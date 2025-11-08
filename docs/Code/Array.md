### Function to Check if Item Exist In Array
```vba
'Function To Check If Item Exist In Array
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error GoTo IsInArrayError:
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function
```
### Function To Sort Array
```vba
Public Sub Test()
    Dim arr() As Variant
    ReDim arr(0 To 3, 0 To 1)
    arr(0, 0) = "One1"
    arr(0, 1) = 1
    
    arr(1, 0) = "Two"
    arr(1, 1) = 2
    
    arr(2, 0) = "Three"
    arr(2, 1) = 3
    
    arr(3, 0) = "One2"
    arr(3, 1) = 1
    
    QuickSortArray arr, , , 1
    
    Dim i As Integer
    For i = 0 To 3
        Debug.Print , arr(i, 0)
    Next
    
End Sub
```
```vba
Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

     
    If IsObject(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)
    
End Sub
```
### Function to Sort Single Dimention Array
```vba
Public Sub QuickSortVector(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next

    'Sort a 1-Dimensional array

    ' SampleUsage: sort arrData
    '
    '   QuickSortVector arrData

    '
    ' Originally posted by Jim Rech 10/20/98 Excel.Programming


    ' Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with an empty variant in the array
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j

        While SortArray(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the item
            varX = SortArray(i)
            SortArray(i) = SortArray(j)
            SortArray(j) = varX

            i = i + 1
            j = j - 1
        End If

    Wend

    If (lngMin < j) Then Call QuickSortVector(SortArray, lngMin, j)
    If (i < lngMax) Then Call QuickSortVector(SortArray, i, lngMax)

End Sub
```
