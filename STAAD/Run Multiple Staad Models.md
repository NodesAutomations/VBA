```vba
Sub simulator()

Dim modelId As Integer
Dim length As Double

Dim objOpenSTAAD As Object
Dim strFileName As String

Dim lNodeNo As Long
Dim lLoadCase As Long
Dim dReactionArray(6) As Double
Dim stat As Integer
For i = 1 To Range("D2").Value

modelId = Cells(4 + i, 2)
length = Cells(4 + i, 3)

Open ThisWorkbook.Path & "\" & modelId & ".std" For Output As #1
Print #1, "STAAD SPACE"
Print #1, "START JOB INFORMATION"
Print #1, "ENGINEER DATE " & Format(Date, "dd-mmm-yy")
Print #1, "END JOB INFORMATION"
Print #1, "INPUT WIDTH 79"
Print #1, "UNIT METER MTON"

Print #1, "JOINT COORDINATES"
Print #1, "1 0 0 0; 2 " & length & " 0 0;"

Print #1, "MEMBER INCIDENCES"
Print #1, "1 1 2;"

Print #1, "DEFINE MATERIAL START"
Print #1, "ISOTROPIC CONCRETE"
Print #1, "E 2.21467e+006"
Print #1, "POISSON 0.17"
Print #1, "DENSITY 2.40262"
Print #1, "ALPHA 1e-005"
Print #1, "DAMP 0.05"
Print #1, "TYPE CONCRETE"
Print #1, "STRENGTH FCU 2812.28"
Print #1, "END DEFINE MATERIAL"


Print #1, "MEMBER PROPERTY AMERICAN"
Print #1, "1 PRIS YD 0.3 ZD 0.3"
Print #1, "CONSTANTS"
Print #1, "MATERIAL CONCRETE ALL"


Print #1, "SUPPORTS"
Print #1, "1 2 FIXED"

Print #1, "LOAD 1 LOADTYPE None  TITLE LOAD CASE 1"
Print #1, "MEMBER LOAD"
Print #1, "1 UNI GY -10"

Print #1, "JOINT LOAD"
Print #1, "1 2 FX -50"
Print #1, "PERFORM ANALYSIS"
Print #1, "FINISH"

Close #1

strFileName = "E:\Documents\Excel sheets\VBA\OpenStaad\" & modelId & ".std"

'Get Staad File
Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
objOpenSTAAD.OpenSTAADFile strFileName
objOpenSTAAD.Analyze

stat = 1
Do While stat = 1
    Application.Wait (Now() + CDate("00:00:02"))  'This method will wait for 2 seconds before checking the analysis status
    stat = objOpenSTAAD.IsAnalyzing()
    SendKeys "{ENTER}", True
Loop

lNodeNo = 1
lLoadCase = 1
objOpenSTAAD.Output.GetSupportReactions lNodeNo, lLoadCase, dReactionArray

Cells(4 + i, 7) = lNodeNo
Cells(4 + i, 8) = dReactionArray(0) / 9.80665
Cells(4 + i, 9) = dReactionArray(1) / 9.80665
Cells(4 + i, 10) = dReactionArray(5) / 9.80665

objOpenSTAAD.CloseSTAADFile
'objOpenSTAAD = Nothing
Next
End Sub
```
