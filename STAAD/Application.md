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

### Check if OpenSTAAD working Properly
```vba
Option Explicit
Sub Main
        Dim objOpenStaad As Object
        Dim stdFile As String
Set objOpenStaad = GetObject(,"StaadPro.OpenSTAAD")
        objOpenStaad.GetSTAADFile stdFile, "TRUE"
        If stdFile="" Then
            MsgBox"Bad"
            Set objOpenStaad = Nothing
            Exit Sub
        End If
MsgBox"Macro Ending"
        Set objOpenStaad = Nothing
    End Sub
    ```
