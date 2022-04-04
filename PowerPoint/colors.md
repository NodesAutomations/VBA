### Get RGB Values of Color Palate
```vba
Private Sub ColorOverride()

Dim pres As Presentation
Dim thm As OfficeTheme
Dim themeColor As themeColor
Dim schemeColors As ThemeColorScheme

Set pres = ActivePresentation

Set schemeColors = pres.Designs(1).SlideMaster.Theme.ThemeColorScheme

    myDark1 = schemeColors(1).RGB         'msoThemeColorDark1
    myLight1 = schemeColors(2).RGB        'msoThemeColorLight
    myDark2 = schemeColors(3).RGB         'msoThemeColorDark2
    myLight2 = schemeColors(4).RGB        'msoThemeColorLight2
    myAccent1 = schemeColors(5).RGB       'msoThemeColorAccent1
    myAccent2 = schemeColors(6).RGB       'msoThemeColorAccent2
    myAccent3 = schemeColors(7).RGB       'msoThemeColorAccent3
    myAccent4 = schemeColors(8).RGB       'msoThemeColorAccent4
    myAccent5 = schemeColors(9).RGB       'msoThemeColorAccent5
    myAccent6 = schemeColors(10).RGB      'msoThemeColorAccent6
    myAccent7 = schemeColors(11).RGB      'msoThemeColorThemeHyperlink
    myAccent8 = schemeColors(12).RGB      'msoThemeColorFollowedHyperlink

    '## THESE LINES RAISE AN ERROR, AS EXPECTED:

    'myAccent9 = schemeColors(13).RGB     
    'myAccent10 = schemeColors(14).RGB
    'myAccent11 = schemeColors(15).RGB
    'myAccent12 = schemeColors(16).RGB

End Sub
```
