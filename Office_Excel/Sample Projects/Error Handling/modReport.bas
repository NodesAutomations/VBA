Attribute VB_Name = "modReport"
Option Explicit

' ADD a type mismatch error to this code using:
' Error 13

' Level 1
' https://excelmacromastery.com/
Public Sub BuildReport()

    On Error GoTo eh
    
    ' Read the user data
    Dim coll As Collection
    Set coll = ReadData(shUSA)

    ' Create the report
    Call WriteData(shReport, coll)
    
    shReport.Activate

done:
    Exit Sub
eh:
    DisplayError Err.Source, Err.Description, "modReport.BuildReport", Erl
End Sub

' Level 2
' https://excelmacromastery.com/
Private Function ReadData(sh As Worksheet) As Collection

    On Error GoTo eh
    
    ' Get the volume setting
    Dim volume As Long
    volume = GetUserSettings
  
    ' Create the collection
    Dim coll As New Collection
    
    ' Get the range
    Dim rg As Range
    Set rg = sh.Range("A1").CurrentRegion

    ' Read through the range
    Dim i As Long
    For i = 2 To rg.Rows.Count
        If CLng(rg.Cells(i, 4).Value) > volume Then
            coll.Add rg.Rows(i).Value
        End If
    Next i

    ' Return the collection
    Set ReadData = coll
     
done:
    Exit Function
eh:
    RaiseError Err.Number, Err.Source, "modReport.ReadData", Err.Description, Erl
End Function

' Level 2
' https://excelmacromastery.com/
Private Sub WriteData(sh As Worksheet, coll As Collection)

    On Error GoTo eh

    ' clear the worksheet
    sh.Cells.ClearContents
    
    ' Get the number of columns from the first range so that
    ' we know the size to output
    Dim columns As Long
    columns = UBound(coll(1), 2)
    
    ' Write out the ranges from the collection
    Dim i As Long
    For i = 1 To coll.Count
        sh.Range("A" & i).Resize(1, columns).Value = coll(i)
    Next i
    
done:
    Exit Sub
eh:
    RaiseError Err.Number, Err.Source, "modReport.WriteData", Err.Description, Erl
End Sub


' Level 3
' https://excelmacromastery.com/
' Description: Reads the user settings from the worksheet
Private Function GetUserSettings() As Long

    On Error GoTo eh
    
    ' Read the user settings from the worksheet
    GetUserSettings = shDashboard.Range("B2").Value
    
done:
    Exit Function
eh:
    RaiseError Err.Number, Err.Source, "modReport.GetUserSettings", Err.Description, Erl
End Function







