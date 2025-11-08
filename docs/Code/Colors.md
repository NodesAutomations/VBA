### function to get Hex Value from shape long color index
```vba
Function GetHexValue(rgbColour As Long) As String
    Dim r As Integer, g As Integer, b As Integer
    Dim rHex As String, gHex As String, bHex As String
  
    r = rgbColour Mod 256
    g = (rgbColour \ 256) Mod 256
    b = (rgbColour \ 256 \ 256) Mod 256
    rHex = hex(r)
    gHex = hex(g)
    bHex = hex(b)
    If Len(rHex) = 1 Then
        rHex = "0" & rHex
    End If
    If Len(gHex) = 1 Then
        gHex = "0" & gHex
    End If
    If Len(bHex) = 1 Then
        bHex = "0" & bHex
    End If
    GetHexValue = rHex & gHex & bHex
End Function
```

### Function to get R,G,B value from shape long color index
```vba
Function GetRGBvalue(rgbColour As Long) As String
    Dim hex As String
    hex = GetHexValue(rgbColour)
    
    Dim r As Integer, g As Integer, b As Integer
    If Len(hex) = 6 Then
        r = Val("&H" & Mid(hex, 1, 2))
        g = Val("&H" & Mid(hex, 3, 2))
        b = Val("&H" & Mid(hex, 5, 2))
      
        GetRGBvalue = r & "," & g & "," & b & ","
    Else
        GetRGBvalue = "0,0,0"
    End If
End Function
```

### Function to convert Hex to RGB
```vba
Function HexToRGB(hex As String) As Long
    Dim r As Integer, g As Integer, b As Integer
    If Len(hex) = 6 Then
        r = Val("&H" & Mid(hex, 1, 2))
        g = Val("&H" & Mid(hex, 3, 2))
        b = Val("&H" & Mid(hex, 5, 2))
        HexToRGB = RGB(r, g, b)
    Else
        HexToRGB = 0
    End If
End Function
```
