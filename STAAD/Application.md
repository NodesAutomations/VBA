### Get Staad Object
```vba
Sub Test()

    Dim objOpenSTAAD As Object
    Dim selbeamsNo As Long
    Dim SelBeams() As Long

    'Launch the OpenSTAAD Object
    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
    
    'Get no. of selected beams
    selbeamsNo = objOpenSTAAD.Geometry.GetNoOfSelectedBeams
    If (selbeamsNo > 0) Then
        .Print selbeamsNo
    End If
End Sub
```
