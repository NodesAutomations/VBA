Attribute VB_Name = "Module1"

Option Explicit

Global Const APPNAME = "xer File Parser & Builder"
Global Const ASC_TAB_CHAR = 9
Global Const EXCEL_ROW_LIMIT = 65536

Global sWorkingFolder   As String
Global sXerFileName     As String
Global iXerFileNum      As Integer

Global bMode            As Boolean 'true=read XER; false=write XER

Global sXerHeader       As String   'stores loaded XER file header

Dim sCurTable           As String
Dim shtCurTable         As Worksheet
Dim shtPrevTable        As Worksheet

Dim lCurTableRow        As Long
Dim lCurTableCol        As Long

Dim lCurTableRowCount   As Long

Dim bStatsOnlyMode      As Boolean

Sub clearSheet(ByVal sSheetName As String)
'clears a specified sheet

Dim sheet As Worksheet

    Set sheet = Worksheets(sSheetName)
    
    If sheet Is Nothing Then
        Exit Sub
    End If
    
    sheet.Cells.Clear
    
    Set sheet = Nothing

End Sub

Sub deleteSheets()
'deletes all worksheets

Dim sheet As Worksheet
Dim bAlerts As Boolean

    Application.StatusBar = "Deleting worksheets..."
    DoEvents
    
    'temporarily disable alerts...
    bAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For Each sheet In Worksheets
    
        Application.StatusBar = "Deleting worksheet " & sheet.Name & "..."
        DoEvents
        
        If UCase(sheet.Name) <> "GENERAL" And UCase(sheet.Name) <> "DIAGNOSTIC" Then
            
            Set sheet = ThisWorkbook.Worksheets(sheet.Name)
            
            If Not (sheet Is Nothing) Then
                sheet.Delete
            End If
        
        End If

    Next
    
    clearSheets
    
    'retore alerts...
    Application.DisplayAlerts = bAlerts
    
    Application.StatusBar = ""
    DoEvents
    
End Sub

Sub clearSheets()
'clears all worksheets

Dim sheet  As Worksheet
    
    Application.StatusBar = "Clearing worksheets..."
    DoEvents
    
    For Each sheet In Worksheets
    
        Application.StatusBar = "Clearing worksheet " & sheet.Name & "..."
        DoEvents
        
        Call clearSheet(sheet.Name)

    Next
    
    Application.StatusBar = ""
    DoEvents
    
End Sub

Sub parseLine(ByVal sLine As String)
'pasrses a line from an XER file

Dim i           As Long
Dim sTemp       As String
Dim bFieldsRow  As Boolean

    If bStatsOnlyMode Then Exit Sub

    bFieldsRow = False

    sTemp = ""
    lCurTableCol = 1
    For i = 1 To Len(sLine)
        
        If Asc(Strings.Mid(sLine, i, 1)) = ASC_TAB_CHAR Then
            
            If sTemp = "%F" Then
                bFieldsRow = True
            End If
            
            If sTemp <> "%F" And sTemp <> "%R" Then
                'shtCurTable.Cells(lCurTableRow, lCurTableCol).NumberFormat = "General"
                shtCurTable.Cells(lCurTableRow, lCurTableCol).NumberFormat = "@"
                shtCurTable.Cells(lCurTableRow, lCurTableCol) = CStr(sTemp)
                If bFieldsRow Then
                    shtCurTable.Cells(lCurTableRow, lCurTableCol).Font.Bold = True
                End If
                sTemp = ""
                lCurTableCol = lCurTableCol + 1
            Else
                sTemp = ""
            End If
            
        Else
        
            sTemp = sTemp & Strings.Mid(sLine, i, 1)
        
        End If
        
    Next i

    If Strings.Trim(sTemp) <> "" Then
    
        'shtCurTable.Cells(lCurTableRow, lCurTableCol).NumberFormat = "General"
        shtCurTable.Cells(lCurTableRow, lCurTableCol).NumberFormat = "@"
    
        shtCurTable.Cells(lCurTableRow, lCurTableCol) = sTemp
        
        If bFieldsRow Then
            shtCurTable.Cells(lCurTableRow, lCurTableCol).Font.Bold = True
        End If
        
    End If

End Sub

Sub buildXerFile()
'builds an XER file from current worksheet contents...

Dim oSheet      As Worksheet
Dim i           As Long
Dim lNumFields  As Long
Dim lSheetRow   As Long
Dim lTotalSheetRows As Long
Dim lSheetCol   As Long
Dim sLine       As String
Dim l           As Long
Dim cTempDate   As Date
Dim sVersion    As String
Dim sOldXerFileName As String

    On Error GoTo buildXerFile_Error
    
    sOldXerFileName = sXerFileName
    
    frmPrompt.chkLoadStatsOnly.Enabled = False
    Call frmPrompt.Show(vbModal)
    
    If (Strings.Trim(sXerFileName) = "") Or (Strings.UCase(sOldXerFileName) = Strings.UCase(sXerFileName)) Then
        sXerFileName = sOldXerFileName
        Exit Sub
    End If
    
    On Error Resume Next
    Kill sXerFileName
    
    'sVersion = ""
    'sVersion = InputBox("Enter version for XER file (e.g., 5.0):", APPNAME, "")
    
    Application.Cursor = xlWait
    DoEvents
    
    iXerFileNum = FreeFile
    Open sXerFileName For Output As #iXerFileNum
    
        'xer header (TP/P3e)...
        'Print #iXerFileNum, "ERMHDR" & Chr(ASC_TAB_CHAR) & sVersion & Chr(ASC_TAB_CHAR) & Format(Now, "yyyy-mm-dd") & Chr(ASC_TAB_CHAR) & "Project" & Chr(ASC_TAB_CHAR) & Application.UserName & Chr(ASC_TAB_CHAR) & Application.OrganizationName & Chr(ASC_TAB_CHAR) & Application.Name & Chr(ASC_TAB_CHAR) & "Project Manager"
        Print #iXerFileNum, sXerHeader
    
        For Each oSheet In Worksheets
        
            If (Strings.UCase(oSheet.Name) <> "GENERAL") And (Strings.UCase(oSheet.Name) <> "DIAGNOSTIC") Then
        
                'make sure sheet has data...
                If (oSheet.Cells(1, 1) <> "") Then
                    
                    If (InStr(1, oSheet.Name, "_") = 0) Then 'make sure this is not a continued table!
                    
                        Application.StatusBar = "Writing Table: " & oSheet.Name & " to file..."
                        DoEvents
                    
                        'write table header...
                        sLine = "%T" & Strings.Chr(ASC_TAB_CHAR) & oSheet.Name
                        Print #iXerFileNum, sLine
                        
                        lNumFields = 0
                        lSheetRow = 1
                        lSheetCol = 1
                        sLine = "%F"
                    
                        'write field header...
                        While Strings.Trim(oSheet.Cells(lSheetRow, lSheetCol)) <> ""
                    
                            lNumFields = lNumFields + 1
                            sLine = sLine & Strings.Chr(ASC_TAB_CHAR) & oSheet.Cells(lSheetRow, lSheetCol)
                            lSheetCol = lSheetCol + 1
                    
                        Wend
                        Print #iXerFileNum, sLine
                        
                    End If
                
                    'write data rows...
                    lTotalSheetRows = getRowCount(oSheet.Name)
                    
                    lSheetRow = 2
                    While Strings.Trim(oSheet.Cells(lSheetRow, 1)) <> ""
                
                        Application.StatusBar = "Writing Table: " & oSheet.Name & " to file (" & CStr(lSheetRow - 1) & " of " & CStr(lTotalSheetRows) & ")..."
                        DoEvents
                        
                        sLine = "%R"
                        For l = 1 To lNumFields
                    
                            If l = 1 Then
                            
                                'do not format id column...
                                sLine = sLine & Strings.Chr(ASC_TAB_CHAR) & oSheet.Cells(lSheetRow, l)
                            
                            Else
                    
                                'is this a date/time field?
                                If InStr(1, oSheet.Cells(1, l), "date") Or InStr(1, oSheet.Cells(1, l), "time") _
                                    Or ((Strings.UCase(oSheet.Name) = "TASKPRED") And ((Strings.UCase(oSheet.Cells(1, l)) = "AREF")) Or ((Strings.UCase(oSheet.Cells(1, l)) = "ARLS"))) Then
                                    sLine = sLine & Strings.Chr(ASC_TAB_CHAR) & Strings.Format(oSheet.Cells(lSheetRow, l), "yyyy-mm-dd hh:mm")
                                'is this a cost field?
                                ElseIf InStr(1, oSheet.Cells(1, l), "cost") Then
                                    sLine = sLine & Strings.Chr(ASC_TAB_CHAR) & Strings.Format(oSheet.Cells(lSheetRow, l), "0.00")
                                Else
                                    sLine = sLine & Strings.Chr(ASC_TAB_CHAR) & oSheet.Cells(lSheetRow, l)
                                End If
                    
                            End If
                    
                        Next l
                
                        Print #iXerFileNum, sLine
                        lSheetRow = lSheetRow + 1
                
                    Wend
                
                End If
            
            End If
            
        Next
    
        'xer footer...
        Print #iXerFileNum, "%E"
    
    Close #iXerFileNum
    
    Application.Cursor = xlDefault
    Application.StatusBar = ""
    DoEvents
    MsgBox "Process complete.", vbInformation, APPNAME

    Exit Sub
    
buildXerFile_Error:
    Application.Cursor = xlDefault
    Application.StatusBar = ""
    DoEvents
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, APPNAME

End Sub

Sub processXerFile()
'loads/parses and XER file into worksheets...

Dim sCurLine        As String

Dim lGeneralRow     As Long
Dim lGeneralCol     As Long

Dim lCurRow         As Long
Dim iCol            As Integer

Dim iNumSheetsForTable  As Integer

    On Error GoTo processXerFile_Error
    
    Call clearSheets
    
    sXerFileName = ""
    frmPrompt.chkLoadStatsOnly.Enabled = True
    Call frmPrompt.Show(vbModal)

    If (Strings.Trim(sXerFileName) = "") Then
        Exit Sub
    End If

    bStatsOnlyMode = (Worksheets("general").Cells(5, 1) <> "")

    If Strings.Trim(sXerFileName) <> "" Then
        
        On Error Resume Next
        Set shtCurTable = Nothing
        Set shtCurTable = Worksheets("Diagnostic")
        If shtCurTable Is Nothing Then
            Set shtCurTable = Worksheets("General")
        End If
        
        On Error GoTo processXerFile_Error
        
        'header...
        lGeneralRow = 6
        Worksheets("general").Cells(lGeneralRow, 1) = "Table:"
        Worksheets("general").Cells(lGeneralRow, 1).Font.Bold = True
        Worksheets("general").Cells(lGeneralRow, 2) = "Row Count:"
        Worksheets("general").Cells(lGeneralRow, 2).Font.Bold = True
        
        lCurTableRowCount = -1
        
        iXerFileNum = FreeFile
        Open sXerFileName For Input As #iXerFileNum
        
            While Not EOF(iXerFileNum)
            
                Line Input #iXerFileNum, sCurLine
                
                'see if header...
                If (Strings.Mid(sCurLine, 1, Len("ERMHDR")) = "ERMHDR") Then
                    sXerHeader = sCurLine   'store xer header
                
                'see if table...
                ElseIf (Strings.Mid(sCurLine, 1, 2) = "%T") Then
                    
                    sCurTable = Strings.Mid(sCurLine, 4)
                    If lCurTableRowCount <> -1 Then
                        Worksheets("General").Cells(lGeneralRow, 2) = lCurTableRowCount
                    End If
                    
                    lGeneralRow = lGeneralRow + 1
                    
                    Worksheets("General").Cells(lGeneralRow, 1) = sCurTable
                    Worksheets("General").Cells(lGeneralRow, 1).Font.Color = vbBlue
                    
                    lCurTableRowCount = 0
                    iNumSheetsForTable = 1
                    
                    Application.StatusBar = "Processing Table: " & sCurTable & "..."
                    DoEvents
                    
                    If Not (bStatsOnlyMode) Then
                    
                        shtCurTable.Columns.AutoFit
                        
                        Set shtPrevTable = shtCurTable
                        Set shtCurTable = Nothing
                        
                        On Error Resume Next
                        Set shtCurTable = Nothing
                        Set shtCurTable = ThisWorkbook.Worksheets(sCurTable)
                        
                        If (shtCurTable Is Nothing) Then
                            Set shtCurTable = ThisWorkbook.Worksheets.Add(, shtPrevTable)
                        End If
                        
                        Call Worksheets("general").Hyperlinks.Add(Worksheets("General").Cells(lGeneralRow, 1), "", sCurTable & "!A1")
                            
                        shtCurTable.Cells.Clear
                        shtCurTable.Cells.ClearContents
                        shtCurTable.Cells.ClearFormats
                        shtCurTable.Cells.NumberFormat = "General"
                        shtCurTable.Name = sCurTable
                        
                        shtCurTable.Activate
                        
                    End If
                    
                    lCurTableRow = 1
                
                'see if fields...
                ElseIf (Strings.Mid(sCurLine, 1, 2) = "%F") Then
                
                    Call parseLine(sCurLine)
                    lCurTableRow = lCurTableRow + 1
                
                'see if row...
                ElseIf (Strings.Mid(sCurLine, 1, 2) = "%R") Then
                
                    lCurTableRowCount = lCurTableRowCount + 1
                    
                    Application.StatusBar = "Processing Table: " & sCurTable & "(" & lCurTableRowCount & ")..."
                    DoEvents
                    
                    'excel limit reached?...
                    If (lCurTableRow >= EXCEL_ROW_LIMIT) And Not (bStatsOnlyMode) Then
                    
                        'cannot (yet) reBuild an XER that required more than
                        'one worksheet to load a single table...
                        Worksheets("General").cmdBuild.Enabled = False
                    
                        Worksheets("General").Cells(lGeneralRow, 2) = lCurTableRowCount
                    
                        iNumSheetsForTable = iNumSheetsForTable + 1
                    
                        lGeneralRow = lGeneralRow + 1
                        Worksheets("General").Cells(lGeneralRow, 1) = "   " & sCurTable & "_" & iNumSheetsForTable
                        Worksheets("General").Cells(lGeneralRow, 1).Font.Color = vbBlue
                        lCurTableRowCount = 0
                        
                        Application.StatusBar = "Processing Table: " & sCurTable & "..."
                        DoEvents
                        
                        If Not (bStatsOnlyMode) Then
                        
                            shtCurTable.Columns.AutoFit
                            
                            Set shtPrevTable = shtCurTable
                            Set shtCurTable = Nothing
                            Set shtCurTable = ThisWorkbook.Worksheets.Add(, shtPrevTable)
                            
                            Call Worksheets("General").Hyperlinks.Add(Worksheets("General").Cells(lGeneralRow, 1), "", sCurTable & "_" & iNumSheetsForTable & "!A1")
                                    
                            shtCurTable.Cells.Clear
                            shtCurTable.Cells.ClearContents
                            shtCurTable.Cells.ClearFormats
                            shtCurTable.Cells.NumberFormat = "General"
                            shtCurTable.Name = sCurTable & "_" & iNumSheetsForTable
                            
                            'copy field names from prevTable to curTable...
                            Call copyFieldHeader(shtPrevTable, shtCurTable)

                        End If

                        lCurTableRow = 2
                        
                    End If
                        
                    Call parseLine(sCurLine)
                    lCurTableRow = lCurTableRow + 1
                    
                End If
                
            Wend
            
            If Not (shtCurTable Is Nothing) And Not (bStatsOnlyMode) Then
                shtCurTable.Columns.AutoFit
            End If
        
        Close #iXerFileNum
        
    End If
    
    If lCurTableRowCount <> -1 Then
        Worksheets("general").Cells(lGeneralRow, 2) = lCurTableRowCount
    End If
    
    Worksheets("general").Columns.AutoFit
    Worksheets("general").Activate
    Worksheets("general").cmdBuild.Enabled = True

    Application.StatusBar = ""
    DoEvents
    
    MsgBox "Process complete.", vbInformation, APPNAME

    Exit Sub

processXerFile_Error:
    Application.StatusBar = ""
    DoEvents
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, APPNAME

End Sub

Sub copyFieldHeader(fromSheet As Worksheet, toSheet As Worksheet)
Dim iCol As Integer

    iCol = 1
    
    While fromSheet.Cells(1, iCol) <> ""
        toSheet.Cells(1, iCol) = fromSheet.Cells(1, iCol)
        toSheet.Cells(1, iCol).Font.Bold = True
        iCol = iCol + 1
    Wend

End Sub


