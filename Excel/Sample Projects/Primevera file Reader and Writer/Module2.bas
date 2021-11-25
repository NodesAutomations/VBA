Attribute VB_Name = "Module2"
Option Explicit

Type rowRec
    sValue As String
    lRow As Long
End Type

Dim lDiagnosticReportRow As Long

Dim lNumRows As Long
Dim aRows() As rowRec

Function getFieldColumn(s As Worksheet, ByVal sName As String) As Integer
Dim iCol As Integer

    iCol = 1
    
    While s.Cells(1, iCol) <> ""
        
        If UCase(s.Cells(1, iCol)) = UCase(sName) Then
        
            getFieldColumn = iCol
            Exit Function
        
        End If

        iCol = iCol + 1
    
    Wend
    
End Function

Function getRowCount(sName As String) As Long
Dim sheet As Worksheet
Dim lRow As Long

    Set sheet = Nothing
    Set sheet = ThisWorkbook.Worksheets("General")

    lRow = 7
    
    While (sheet.Cells(lRow, 1) <> "")
    
        If Strings.UCase(Strings.Trim(sheet.Cells(lRow, 1))) = Strings.UCase(sName) Then
            
            If (sheet.Cells(lRow, 2) <> "") Then
                getRowCount = CLng(sheet.Cells(lRow, 2))
            Else
                getRowCount = 0
            End If
            
            Exit Function
        
        End If
        
        lRow = lRow + 1
    
    Wend
    
    Set sheet = Nothing
    
End Function

Sub appendDiagnosticReport(ByVal sText As String)
Dim sheet As Worksheet
Dim lRow As Long

    Set sheet = Nothing
    Set sheet = ThisWorkbook.Worksheets("Diagnostic")

    sheet.Cells(lDiagnosticReportRow, 1) = sText
    lDiagnosticReportRow = lDiagnosticReportRow + 1

    Set sheet = Nothing

End Sub

Public Sub SortArray(ByRef Arr() As rowRec, ByVal ascending As Boolean)
Dim l As Long
Dim r As Long

    l = 0
    r = UBound(Arr)

    If (ascending) Then
        Call QuickSort(Arr, l, r, 1)
    Else
        Call QuickSort(Arr, l, r, -1)
    End If
    
End Sub

Private Sub QuickSort(ByRef Arr() As rowRec, ByVal l As Long, ByVal r As Long, ByVal flag As Integer)
Dim i As Long
Dim j As Long
Dim temp As rowRec
Dim ret As Long

    
    If (r <= l) Then Exit Sub
    
    i = l - 1
    j = r

    Do While (True)
    
        Do
            i = i + 1
            ret = StrComp(Arr(i).sValue, Arr(r).sValue)
            
            If ret = 0 Then
            
                If Arr(i).lRow < Arr(r).lRow Then
                    ret = -1
                ElseIf Arr(i).lRow > Arr(r).lRow Then
                    ret = 1
                Else
                    ret = 0
                End If
            
            End If
            
            ret = ret * flag
        
        Loop While (ret < 0)


        Do While (j > 0)
            j = j - 1
            ret = StrComp(Arr(j).sValue, Arr(r).sValue)
            
            If ret = 0 Then
            
                If Arr(j).lRow < Arr(r).lRow Then
                    ret = -1
                ElseIf Arr(j).lRow > Arr(r).lRow Then
                    ret = 1
                Else
                    ret = 0
                End If
                
            End If
            
            ret = ret * flag
            If (ret <= 0) Then Exit Do
        Loop
        
        If (i > j) Then Exit Do
        temp.lRow = Arr(i).lRow
        temp.sValue = Arr(i).sValue
        
        Arr(i).lRow = Arr(j).lRow
        Arr(i).sValue = Arr(j).sValue
    
        Arr(j).lRow = temp.lRow
        Arr(j).sValue = temp.sValue
        
    Loop
    
    temp.lRow = Arr(i).lRow
    temp.sValue = Arr(i).sValue
        
    Arr(i).lRow = Arr(r).lRow
    Arr(i).sValue = Arr(r).sValue
    
    Arr(r).lRow = temp.lRow
    Arr(r).sValue = temp.sValue
    
    Call QuickSort(Arr, l, i - 1, flag)
    Call QuickSort(Arr, i + 1, r, flag)
    
End Sub

Function getRowsText(Arr() As Long) As String
Dim l As Long

    getRowsText = ""
    
    For l = 0 To UBound(Arr)
    
        getRowsText = getRowsText & Arr(l)
        
        If (l <> UBound(Arr)) Then
            getRowsText = getRowsText & "; "
        End If
        
    Next l

End Function

Sub bubbleSort(ByRef aList() As Long)
Dim i As Long
Dim j As Long
Dim lTemp As Long
Dim lPass As Long
        
    lPass = 0
        
    For i = UBound(aList) - 1 To 0 Step -1

        lPass = lPass + 1
        Application.StatusBar = "Sorting list (pass " & CStr(lPass) & " of " & CStr(UBound(aList)) & ")..."
        DoEvents

        For j = 0 To i

            If aList(j) > aList(j + 1) Then

                lTemp = aList(j)

                aList(j) = aList(j + 1)

                aList(j + 1) = lTemp

            End If

        Next j

    Next i

End Sub

Function binarySearch(aList() As Long, lFind As Long) As Boolean
Dim lLow As Long
Dim lHigh As Long
Dim lMiddle As Long

    binarySearch = False

    lLow = 0
    lHigh = UBound(aList)

    While (lLow <= lHigh)
    
        lMiddle = (lLow + lHigh) / 2
        
        If (lFind = aList(lMiddle)) Then
            binarySearch = True
            Exit Function
        ElseIf (lFind < aList(lMiddle)) Then
            lHigh = lMiddle - 1      'search low end of array
        Else
            lLow = lMiddle + 1       'search high end of array
        End If
    
    Wend

End Function

Sub crossCheckFK(ByVal sTable1 As String, ByVal sField1 As String, ByVal sTable2 As String, ByVal sField2 As String)
Dim Sheet1 As Worksheet
Dim Sheet2 As Worksheet
Dim iFieldCol1 As Integer
Dim iFieldCol2 As Integer
Dim lRowCount1 As Long
Dim lRowCount2 As Long
Dim aPKs() As Long
Dim aFKs() As Long
Dim aFKRows() As Long
Dim lNumPKs As Long
Dim lNumFKs As Long
Dim isUnique As Boolean
Dim bFound As Boolean
Dim lFK As Long
Dim l As Long
Dim m As Long
Dim sUdfTypeId As String
Dim sUdfTypeTable As String
Dim bSkip As Boolean

    Application.Cursor = xlWait
    DoEvents

    Call appendDiagnosticReport("")
    Call appendDiagnosticReport("BEGIN DIAGNOSTIC - cross check " & sTable1 & ":" & sField1 & " / " & sTable2 & ":" & sField2)
    
    Set Sheet1 = ThisWorkbook.Worksheets(sTable1)
    Set Sheet2 = ThisWorkbook.Worksheets(sTable2)
    
    'find field column...
    iFieldCol1 = getFieldColumn(Sheet1, sField1)
    iFieldCol2 = getFieldColumn(Sheet2, sField2)
    
    'get the loaded row count sheet/table...
    lRowCount1 = getRowCount(sTable1)
    lRowCount2 = getRowCount(sTable2)
    
    'obtain a unique list of values from table2...
    Application.StatusBar = "Generating unique FK list..."
    DoEvents
    
    ReDim aFKs(0)
    ReDim aFKRows(0)
    lNumFKs = -1
    
    For l = 2 To lRowCount2 + 1
        
        Application.StatusBar = "Loading FK list : processing (" & CStr(l) & " of " & CStr(lRowCount2) & ")..."
        DoEvents
        
        bSkip = False
        
        'UDFVALUE FK check (filter)...
        If (VBA.InStr(1, UCase(sTable2), "UDFVALUE", vbTextCompare) <> 0) Then
            sUdfTypeId = ThisWorkbook.Worksheets(sTable2).Cells(l, 1)
            sUdfTypeTable = getUdfTypeTable(sUdfTypeId)
            bSkip = (UCase(sUdfTypeTable) <> UCase(sTable1))
        End If
        
        If Not (bSkip) Then
        
            lFK = CLng(Sheet2.Cells(l, iFieldCol2))
            
            'isUnique = True
            
            'For m = 0 To UBound(aFKs)
            '    If aFKs(m) = lFK Then
            '        isUnique = False
            '        Exit For
            '    End If
            'Next m
        
            'If isUnique Then
                lNumFKs = lNumFKs + 1
                ReDim Preserve aFKs(lNumFKs)
                aFKs(lNumFKs) = lFK
                ReDim Preserve aFKRows(lNumFKs)
                aFKRows(lNumFKs) = l
                
            'End If
        
        End If
        
    Next l
    
    'load PKs...
    Application.StatusBar = "Loading PK list..."
    DoEvents
    
    ReDim aPKs(0)
    lNumPKs = -1
    For m = 2 To lRowCount1 + 1
        lNumPKs = lNumPKs + 1
        ReDim Preserve aPKs(lNumPKs)
        aPKs(lNumPKs) = CLng(Sheet1.Cells(m, iFieldCol1))
    Next m
    
    'sort PKs...
    Call bubbleSort(aPKs)
    
    'cross check (PKs / FKs)...
    For l = 0 To UBound(aFKs)
        
        If ((aFKs(l)) = 0) Then GoTo skip
        
        Application.StatusBar = "cross checking: " & aFKs(l) & "  (" & CStr(l + 1) & " of " & CStr(UBound(aFKs)) & ")..."
        DoEvents
        
        'OLD CODE...
        'bFound = False
        'For m = 2 To lRowCount1 + 1
        '    If Sheet1.Cells(m, iFieldCol1) = aFks(l) Then
        '        bFound = True
        '        Exit For
        '    End If
        'Next m
        
        'NEW CODE...
        bFound = binarySearch(aPKs, aFKs(l))
        
        If Not (bFound) Then
            Call appendDiagnosticReport("      cross check failed for: " & sField2 & ": " & CStr(aFKs(l)) & "  (row #: " & CStr(aFKRows(l)) & " in table/worksheet: " & sTable2 & ")")
        End If
skip:

    Next l
    
    Call appendDiagnosticReport("END DIAGNOSTIC - cross check FK")

    Application.Cursor = xlDefault
    DoEvents
    
    Application.StatusBar = ""
    DoEvents
    
End Sub

Sub checkForDuplicate(ByVal sTable As String, aFields() As String)
Dim sheet As Worksheet
Dim sFieldName As String
Dim iFieldCol As Integer
Dim lRowCount As Long
Dim l As Long
Dim i As Long
Dim sCurValue As String

Dim lNumDupRows As Long
Dim aDupRows() As Long

    Call appendDiagnosticReport("")
    Call appendDiagnosticReport("BEGIN DIAGNOSTIC - check for Duplicates")

    Set sheet = ThisWorkbook.Worksheets(sTable)
    
    For i = 0 To UBound(aFields)
       
        sFieldName = aFields(i)
            
        Call appendDiagnosticReport("   checking " & sFieldName & " for Duplicates...")
            
        'find field column...
        iFieldCol = getFieldColumn(sheet, sFieldName)
         
        'get the loaded row count for this sheet/table...
        lRowCount = getRowCount(sheet.Name)
         
        lNumRows = -1
        ReDim aRows(0)
        For l = 2 To lRowCount + 1
        
            lNumRows = lNumRows + 1
            ReDim Preserve aRows(lNumRows)
            aRows(lNumRows).lRow = l
            aRows(lNumRows).sValue = sheet.Cells(l, iFieldCol)
        
        Next l
        
        'sort array...
        Call SortArray(aRows, True)
        
        sCurValue = ""
        
        lNumDupRows = 0
        ReDim aDupRows(0)
        
        For l = 0 To UBound(aRows)
        
            If (aRows(l).sValue <> sCurValue) Then
                
                If UBound(aDupRows) > 0 Then
                    Call appendDiagnosticReport("      duplicate value '" & sCurValue & "' found at row(s): " & getRowsText(aDupRows))
                End If
                
                sCurValue = aRows(l).sValue
                
                lNumDupRows = 0
                ReDim aDupRows(0)
                aDupRows(0) = aRows(l).lRow
                
            ElseIf (aRows(l).sValue = sCurValue) Then
            
                lNumDupRows = lNumDupRows + 1
                ReDim Preserve aDupRows(lNumDupRows)
                aDupRows(lNumDupRows) = aRows(l).lRow
            
            End If
        
        Next l
        
        If UBound(aDupRows) > 0 Then
            Call appendDiagnosticReport("      duplicate value '" & sCurValue & "' found at rows: " & getRowsText(aDupRows))
        End If
        
    Next i
    
    Call appendDiagnosticReport("END DIAGNOSTIC - check for Duplciates")
    
End Sub

Sub checkForNull(ByVal sTable As String, aFields() As String)
Dim sheet As Worksheet
Dim sFieldName As String
Dim iFieldCol As Integer
Dim lRowCount As Long
Dim lRow As Long
Dim i As Long
Dim lNumNullRows As Long
Dim aNullRows() As Long

    Application.Cursor = xlWait
    DoEvents

    Call appendDiagnosticReport("")
    Call appendDiagnosticReport("BEGIN DIAGNOSTIC - check for NULL")

    Set sheet = ThisWorkbook.Worksheets(sTable)
    
    For i = 0 To UBound(aFields)
    
        sFieldName = aFields(i)
            
        Call appendDiagnosticReport("   checking " & sFieldName & " for NULL...")
        Application.StatusBar = "checking " & sFieldName & " for NULL..."
        DoEvents
            
        'find field column...
        iFieldCol = getFieldColumn(sheet, sFieldName)
         
        'get the loaded row count for this sheet/table...
        lRowCount = getRowCount(sheet.Name)
         
        lNumNullRows = -1
        ReDim aNullRows(0)
        aNullRows(0) = -1
         
        For lRow = 2 To lRowCount + 1
                 
            If sheet.Cells(lRow, iFieldCol) = "" Then
                
                lNumNullRows = lNumNullRows + 1
                ReDim Preserve aNullRows(lNumNullRows)
                aNullRows(lNumNullRows) = lRow
                
            End If
                 
        Next lRow
        
        If (aNullRows(0) <> -1) Then
            Call appendDiagnosticReport("      NULL value found at row(s): " & getRowsText(aNullRows))
        End If
    
    Next i
    
    Call appendDiagnosticReport("END DIAGNOSTIC - Check for NULL")
    
    Application.Cursor = xlDefault
    DoEvents
    
    Application.StatusBar = ""
    DoEvents
    
End Sub
Sub formatDiagnosticSheet()
Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = Nothing
    Set sheet = ThisWorkbook.Worksheets("Diagnostic")
    
    sheet.Columns.AutoFit
    
End Sub

Sub initDiagnosticSheet()
Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = Nothing
    Set sheet = ThisWorkbook.Worksheets("Diagnostic")
                        
    If (sheet Is Nothing) Then
        Set sheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets("General"))
        sheet.Name = "Diagnostic"
    End If
    sheet.Cells.Clear
    sheet.Activate
    Call sheet.Cells(1, 1).Activate
    
    lDiagnosticReportRow = 1
    Call appendDiagnosticReport("DIAGNOSTIC for: " & sXerFileName)
    Call appendDiagnosticReport("")

End Sub

Function getUdfTypeTable(sUdfTypeId As String) As String
Dim lRow As Long

    getUdfTypeTable = ""
    
    lRow = 2
    
    While ThisWorkbook.Worksheets("UDFTYPE").Cells(lRow, 1) <> ""
    
        If ThisWorkbook.Worksheets("UDFTYPE").Cells(lRow, 1) = CStr(sUdfTypeId) Then
            getUdfTypeTable = ThisWorkbook.Worksheets("UDFTYPE").Cells(lRow, 2)
            Exit Function
        End If
    
        lRow = lRow + 1
    
    Wend

End Function
