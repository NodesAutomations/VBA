### Get Reaction from STAAD

```vba
Sub Reaction()
Dim objOpenSTAAD As Object
Dim lNodeNo As Long
Dim lLoadCase As Long
Dim dReactionArray(6) As Double
'Get the application object
Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
lNodeNo = 1
lLoadCase = 1
objOpenSTAAD.Output.GetSupportReactions lNodeNo, lLoadCase, dReactionArray

Debug.Print dReactionArray(0) / 9.80665
Debug.Print dReactionArray(1) / 9.80665
Debug.Print dReactionArray(2) / 9.80665

Debug.Print dReactionArray(3) / 9.80665
Debug.Print dReactionArray(4) / 9.80665
Debug.Print dReactionArray(5) / 9.80665
End Sub
```
### Get Beam Section forces
```vba
Sub Main()
Dim StartTime As Double
Dim SecondsElapsed As Double

'Start
StartTime = Timer
Dim objOpenSTAAD As Object
Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
Dim lBeamNo As Long
Dim lLoadCase As Long
Dim lForcesArray(6) As Double
Dim dDistance As Double
Dim dBeamLength As Double
Dim strResults As String
Dim strFileFolder As String

objOpenSTAAD.GetSTAADFileFolder strFileFolder

If Len(Dir(strFileFolder & "\STD_Results", vbDirectory)) = 0 Then
 MkDir strFileFolder & "\STD_Results"
End If
Open strFileFolder & "\STD_Results\Moments.csv" For Output As #1
'Beams
For i =1 To 300
lBeamNo=i
	'LoadCase
	For j =100 To 790
	lLoadCase=j
		'Points
		For k =0 To 12
		dBeamLength = objOpenSTAAD.Geometry.GetBeamLength(lBeamNo)
		dDistance = k * dBeamLength
		objOpenSTAAD.Output.GetIntermediateMemberForcesAtDistance (lBeamNo, dDistance, lLoadCase,lForcesArray)
		Print #1,i &  "," & j &  "," & k &  ","& lForcesArray(0) & "," & lForcesArray(1)& "," &lForcesArray(2)&  ","& lForcesArray(3) & "," & lForcesArray(4)& "," &lForcesArray(5)
		Next	
	Next
Next

For i =333 To 684
lBeamNo=i
	'LoadCase
	For j =100 To 790
	lLoadCase=j
		'Points
		For k =0 To 12
		dBeamLength = objOpenSTAAD.Geometry.GetBeamLength(lBeamNo)
		dDistance = k * dBeamLength
		objOpenSTAAD.Output.GetIntermediateMemberForcesAtDistance (lBeamNo, dDistance, lLoadCase,lForcesArray)
		Print #1,i &  "," & j &  "," & k &  ","& lForcesArray(0) & "," & lForcesArray(1)& "," &lForcesArray(2)&  ","& lForcesArray(3) & "," & lForcesArray(4)& "," &lForcesArray(5)
		Next	
	Next
Next
Close #1
SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox "code ran successfully in " & SecondsElapsed &  " seconds",vbInformation
End Sub
```
