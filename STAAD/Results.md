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
