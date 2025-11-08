Attribute VB_Name = "DynamicRibbon"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
#End If
Public Fnd As String
Public Rplc As String

Public RefreshRibbon As IRibbonUI
Public EditboxText As String
Public ComboboxText As String
Public ComboItemCount As Long
Public Dropdown As String
Public DropdownItemCount As Long
Public DropdownSelectedItem As Long
Public ChkBx(1 To 6) As Boolean
Public Tglbtn(1 To 6) As Boolean

Public Sub RefreshControls(ribbon As IRibbonUI)
    Set RefreshRibbon = ribbon                   ' Set Ribbon onLoad

    saveGlobal RefreshRibbon, "RibbonPtr"        'This Function to Save and ReStore Ribbon after Replacing Below Items or any Fault
    ' Contnue Replacing to save values of Ribbon Controls Using:   Sub VBRplcr(PrcName As String, Fnd As String, Rplc As String)'

        EditboxText = "Day"                      ' EditboxText1 Text value

        ''''''''''''''''''''''''
        ComboboxText = "AAA"                     ' Combobox1 Text value
        ComboItemCount = 6                       '           Itmes Count

        '''
        Dropdown = "Friday"                      ' Dropdown1: Text value
        DropdownItemCount = 6                    '            Itmes Count
        DropdownSelectedItem = 5                 '            Itme Number

        '''
        ChkBx(1) = True                          'Free select (1 to 3)
        ChkBx(2) = True
        ChkBx(3) = True
        '''
        ChkBx(4) = False                         'One selected Option From Group select (4 to 6)
        ChkBx(5) = True
        ChkBx(6) = False

        Tglbtn(1) = False                        'Free select (1 to 3)
        Tglbtn(2) = True
        Tglbtn(3) = False
        '''
        Tglbtn(4) = False                        'One selected Option From Group select (4 to 6)
        Tglbtn(5) = False
        Tglbtn(6) = True


    End Sub

Public Sub Editbox_getText(control As IRibbonControl, ByRef returnedVal)
    If control.id = "Editbox1" Then
        returnedVal = EditboxText
    End If
End Sub

Public Sub Editbox_onChange(control As IRibbonControl, Text As String)
    EditboxText = "Day"

    Fnd = ""
    Fnd = "EditboxText = " & """" & EditboxText & """"
    Rplc = ""
    Rplc = "EditboxText = " & """" & Text & """"
    VBRplcr "RefreshControls", Fnd, Rplc
    VBRplcr "Editbox_getText", Fnd, Rplc
    VBRplcr "Editbox_onChange", Fnd, Rplc
    If control.id = "Editbox1" Then

        EditboxText = Text
    End If
    If RefreshRibbon Is Nothing Then Set RefreshRibbon = GetGlobal("RibbonPtr")
    RefreshRibbon.Invalidate
End Sub

Public Sub Combobox_getText(control As IRibbonControl, ByRef returnedVal)
    If control.id = "Combobox1" Then
        returnedVal = ComboboxText
    End If

End Sub

Public Sub Combobox_onChange(control As IRibbonControl, Text As String)

    ComboboxText = "AAA"

    If control.id = "Combobox1" Then
        Fnd = ""
        Fnd = "ComboboxText = " & """" & ComboboxText & """"
        Rplc = ""
        Rplc = "ComboboxText = " & """" & Text & """"
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Combobox_getText", Fnd, Rplc
        VBRplcr "Combobox_onChange", Fnd, Rplc

        ComboboxText = Text

    End If
    ''''''''''''''''''''''''''''''''''''''
    If RefreshRibbon Is Nothing Then Set RefreshRibbon = GetGlobal("RibbonPtr")
    RefreshRibbon.Invalidate
End Sub

Public Sub Combobox_getItemCount(control As IRibbonControl, ByRef returnedVal)

    If control.id = "Combobox1" Then
        returnedVal = 6
    End If
End Sub

Public Sub ComboboxgetItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Dim ComboItemLabel As Variant
    If control.id = "Combobox1" Then
        ComboItemLabel = Array("AAA", "BBB", "CCC", "DDD", "EEE", "FFF")

        Dim I As Long

        returnedVal = ComboItemLabel(index)
    Else

    End If

End Sub

Public Sub Dropdown_getItemCount(control As IRibbonControl, ByRef returnedVal)

    DropdownItemCount = 6

    If control.id = "Dropdown1" Then
        returnedVal = DropdownItemCount

    End If
End Sub

Public Sub Dropdown_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    DropdownSelectedItem = index
    returnedVal = WeekdayName(index + 1)
End Sub

Public Sub Dropdown_getSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)

    DropdownSelectedItem = 5
    returnedVal = DropdownSelectedItem
End Sub

Public Sub GetAction(control As IRibbonControl, id As String, index As Integer)

    If control.id = "Dropdown1" Then
        Dropdown = "Friday"
        DropdownSelectedItem = 5

        Fnd = "": Rplc = ""
        Fnd = "Dropdown = " & """" & Dropdown & """"
        Rplc = "Dropdown = " & """" & WeekdayName(index + 1) & """"
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "GetAction", Fnd, Rplc

        Fnd = "": Rplc = ""
        Fnd = "DropdownItemCount = " & DropdownItemCount
        Rplc = "DropdownItemCount = " & DropdownItemCount
        VBRplcr "RefreshControls", Fnd, Rplc

        Fnd = ""
        Fnd = "DropdownSelectedItem = " & DropdownSelectedItem
        Rplc = ""
        Rplc = "DropdownSelectedItem = " & index
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Dropdown_getSelectedItemIndex", Fnd, Rplc
        VBRplcr "GetAction", Fnd, Rplc
        '''''''''Your Action


    ElseIf control.id = "Dropdown2" Then

    ElseIf control.id = "Dropdown3" Then

    End If
    If RefreshRibbon Is Nothing Then Set RefreshRibbon = GetGlobal("RibbonPtr")
    RefreshRibbon.Invalidate

End Sub

Public Sub Checkbox_getPressed(control As IRibbonControl, ByRef returnedVal)

    ChkBx(1) = True
    ChkBx(2) = True
    ChkBx(3) = True
    ChkBx(4) = False
    ChkBx(5) = True
    ChkBx(6) = False


    If control.id = "Checkbox1" Then
        returnedVal = ChkBx(1)
    ElseIf control.id = "Checkbox2" Then
        returnedVal = ChkBx(2)
    ElseIf control.id = "Checkbox3" Then
        returnedVal = ChkBx(3)

    ElseIf control.id = "Checkbox4" Then
        returnedVal = ChkBx(4)
    ElseIf control.id = "Checkbox5" Then
        returnedVal = ChkBx(5)
    ElseIf control.id = "Checkbox6" Then
        returnedVal = ChkBx(6)
    End If
    Exit Sub

End Sub

Public Sub Checkbox_onAction(control As IRibbonControl, pressed As Boolean)

    Fnd = "": Rplc = ""
    If control.id = "Checkbox1" Then
        Fnd = "ChkBx(1) = " & ChkBx(1)
        Rplc = "ChkBx(1) = " & pressed
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Checkbox_getPressed", Fnd, Rplc

        ChkBx(1) = pressed
        'You Action Here

    ElseIf control.id = "Checkbox2" Then
        Fnd = "ChkBx(2) = " & ChkBx(2)
        Rplc = "ChkBx(2) = " & pressed
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Checkbox_getPressed", Fnd, Rplc

        ChkBx(2) = pressed
        'You Action Here

    ElseIf control.id = "Checkbox3" Then
        Fnd = "ChkBx(3) = " & ChkBx(3)
        Rplc = "ChkBx(3) = " & pressed
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Checkbox_getPressed", Fnd, Rplc

        ChkBx(3) = pressed
        'You Action Here

    ElseIf control.id = "Checkbox4" Then
        If pressed = True Then
            ChkBx(4) = pressed
            ChkBx(5) = Not pressed
            ChkBx(6) = Not pressed
            Fnd = "ChkBx(4) = " & Not pressed: Rplc = "ChkBx(4) = " & pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc

            Fnd = "ChkBx(5) = " & pressed: Rplc = "ChkBx(5) = " & Not pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc

            Fnd = "ChkBx(6) = " & pressed: Rplc = "ChkBx(6) = " & Not pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc
            'You Action Here

        End If
    ElseIf control.id = "Checkbox5" Then
        If pressed = True Then
            ChkBx(5) = pressed
            ChkBx(4) = Not pressed
            ChkBx(6) = Not pressed

            Fnd = "ChkBx(5) = " & Not pressed: Rplc = "ChkBx(5) = " & pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc

            Fnd = "ChkBx(4) = " & pressed: Rplc = "ChkBx(4) = " & Not pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc

            Fnd = "ChkBx(6) = " & pressed: Rplc = "ChkBx(6) = " & Not pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc
            'You Action Here

        End If
    ElseIf control.id = "Checkbox6" Then
        If pressed = True Then
            ChkBx(6) = pressed
            ChkBx(4) = Not pressed
            ChkBx(5) = Not pressed

            Fnd = "ChkBx(6) = " & Not pressed: Rplc = "ChkBx(6) = " & pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc

            Fnd = "ChkBx(4) = " & pressed: Rplc = "ChkBx(4) = " & Not pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc

            Fnd = "ChkBx(5) = " & pressed: Rplc = "ChkBx(5) = " & Not pressed
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Checkbox_getPressed", Fnd, Rplc
            'You Action Here

        End If
    End If

    If RefreshRibbon Is Nothing Then Set RefreshRibbon = GetGlobal("RibbonPtr")
    RefreshRibbon.Invalidate

End Sub

Public Sub Togglebutton_getLabel(control As IRibbonControl, ByRef returnedVal)

    Tglbtn(1) = False
    Tglbtn(2) = True
    Tglbtn(3) = False
    Tglbtn(4) = False
    Tglbtn(5) = False
    Tglbtn(6) = True

    If control.id = "Togglebutton1" Then
        If Tglbtn(1) = True Then
            returnedVal = "On"
        Else
            returnedVal = "Off"
        End If
    ElseIf control.id = "Togglebutton2" Then
        If Tglbtn(2) = True Then
            returnedVal = "On"
        Else
            returnedVal = "Off"
        End If
    ElseIf control.id = "Togglebutton3" Then
        If Tglbtn(3) = True Then
            returnedVal = "On"
        Else
            returnedVal = "Off"
        End If
    ElseIf control.id = "Togglebutton4" Then
        If Tglbtn(4) = False Then
            returnedVal = "Off"
        Else
            returnedVal = "On"
        End If
    ElseIf control.id = "Togglebutton5" Then
        If Tglbtn(5) = False Then
            returnedVal = "Off"
        Else
            returnedVal = "On"
        End If
    ElseIf control.id = "Togglebutton6" Then
        If Tglbtn(6) = False Then
            returnedVal = "Off"
        Else
            returnedVal = "On"
        End If
    End If
End Sub

Public Sub Togglebutton_getPressed(control As IRibbonControl, ByRef returnedVal)

    Tglbtn(1) = False
    Tglbtn(2) = True
    Tglbtn(3) = False
    Tglbtn(4) = False
    Tglbtn(5) = False
    Tglbtn(6) = True
    If control.id = "Togglebutton1" Then
        returnedVal = Tglbtn(1)
    ElseIf control.id = "Togglebutton2" Then
        returnedVal = Tglbtn(2)
    ElseIf control.id = "Togglebutton3" Then
        returnedVal = Tglbtn(3)

    ElseIf control.id = "Togglebutton4" Then
        returnedVal = Tglbtn(4)
    ElseIf control.id = "Togglebutton5" Then
        returnedVal = Tglbtn(5)
    ElseIf control.id = "Togglebutton6" Then
        returnedVal = Tglbtn(6)
    End If
    Exit Sub
End Sub

Public Sub Togglebutton_onAction(control As IRibbonControl, ByRef cancelDefault)

    Fnd = "": Rplc = ""
    If control.id = "Togglebutton1" Then
        Fnd = "Tglbtn(1) = " & Tglbtn(1)
        Rplc = "Tglbtn(1) = " & cancelDefault
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Togglebutton_getPressed", Fnd, Rplc
        VBRplcr "Togglebutton_getLabel", Fnd, Rplc
        Tglbtn(1) = cancelDefault
        'You Action Here


    ElseIf control.id = "Togglebutton2" Then
        Fnd = "Tglbtn(2) = " & Tglbtn(2)
        Rplc = "Tglbtn(2) = " & cancelDefault
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Togglebutton_getPressed", Fnd, Rplc
        VBRplcr "Togglebutton_getLabel", Fnd, Rplc
        Tglbtn(2) = cancelDefault
        'You Action Here


    ElseIf control.id = "Togglebutton3" Then
        Fnd = "Tglbtn(3) = " & Tglbtn(3)
        Rplc = "Tglbtn(3) = " & cancelDefault
        VBRplcr "RefreshControls", Fnd, Rplc
        VBRplcr "Togglebutton_getPressed", Fnd, Rplc
        VBRplcr "Togglebutton_getLabel", Fnd, Rplc
        Tglbtn(3) = cancelDefault
        'You Action Here


    ElseIf control.id = "Togglebutton4" Then
        If cancelDefault = True Then
            Tglbtn(4) = cancelDefault
            Tglbtn(5) = Not cancelDefault
            Tglbtn(6) = Not cancelDefault
            Fnd = "Tglbtn(4) = " & Not cancelDefault: Rplc = "Tglbtn(4) = " & cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc

            Fnd = "Tglbtn(5) = " & cancelDefault: Rplc = "Tglbtn(5) = " & Not cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc

            Fnd = "Tglbtn(6) = " & cancelDefault: Rplc = "Tglbtn(6) = " & Not cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc
            'You Action Here



        End If
    ElseIf control.id = "Togglebutton5" Then
        If cancelDefault = True Then
            Tglbtn(5) = cancelDefault
            Tglbtn(4) = Not cancelDefault
            Tglbtn(6) = Not cancelDefault

            Fnd = "Tglbtn(5) = " & Not cancelDefault: Rplc = "Tglbtn(5) = " & cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc

            Fnd = "Tglbtn(4) = " & cancelDefault: Rplc = "Tglbtn(4) = " & Not cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc

            Fnd = "Tglbtn(6) = " & cancelDefault: Rplc = "Tglbtn(6) = " & Not cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc
            'You Action Here

        End If
    ElseIf control.id = "Togglebutton6" Then
        If cancelDefault = True Then
            Tglbtn(6) = cancelDefault
            Tglbtn(4) = Not cancelDefault
            Tglbtn(5) = Not cancelDefault

            Fnd = "Tglbtn(6) = " & Not cancelDefault: Rplc = "Tglbtn(6) = " & cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc

            Fnd = "Tglbtn(4) = " & cancelDefault: Rplc = "Tglbtn(4) = " & Not cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc

            Fnd = "Tglbtn(5) = " & cancelDefault: Rplc = "Tglbtn(5) = " & Not cancelDefault
            VBRplcr "RefreshControls", Fnd, Rplc
            VBRplcr "Togglebutton_getPressed", Fnd, Rplc
            VBRplcr "Togglebutton_getLabel", Fnd, Rplc
            'You Action Here

        End If
    End If
    If RefreshRibbon Is Nothing Then Set RefreshRibbon = GetGlobal("RibbonPtr")
    RefreshRibbon.Invalidate

End Sub

Public Sub saveGlobal(Glbl As Object, GlblName As String)

    #If VBA7 Then
        Dim lngRibPtr As LongPtr
    #Else
        Dim lngRibPtr As Long
    #End If
    lngRibPtr = ObjPtr(Glbl)
    With ThisWorkbook
        On Error Resume Next
        .Names(GlblName).Delete
        On Error GoTo 0
        .Names.Add GlblName, lngRibPtr
        .Saved = True
    End With
End Sub

Public Function GetGlobal(GlblName As String) As Object

    #If VBA7 Then
        Dim X As LongPtr
        X = CLngPtr(Mid(ThisWorkbook.Names(GlblName).RefersTo, 2))
    #Else
        Dim X As Long
        X = CLng(Mid(ThisWorkbook.Names(GlblName).RefersTo, 2))
    #End If
    Dim objRibbon As Object
    CopyMemory objRibbon, X, Len(X)
    Set GetGlobal = objRibbon
End Function

Sub VBRplcr(PrcName As String, Fnd As String, Rplc As String)
    'Microsoft Visual Basic for Applications Extensibility 5.3 is required
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim ThisLine As String
    Dim N As Long
    Dim ProcStrLn As Long, ProcAcStrLn As Long, ProcCntLn As Long, PrcCnountLine As Long
    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        With VBComp
            If .Type = vbext_ct_StdModule Then
                With .CodeModule
                    If InStr(1, .Lines(1, .CountOfLines), PrcName) > 0 Then
                        On Error Resume Next
                        ProcStrLn = .ProcStartLine(PrcName, vbext_pk_Proc)
                        ProcAcStrLn = .ProcBodyLine(PrcName, vbext_pk_Proc)
                        ProcCntLn = .ProcCountLines(PrcName, vbext_pk_Proc)
                        PrcCnountLine = ProcCntLn - (ProcAcStrLn - ProcStrLn)
                        If PrcName = .ProcOfLine(ProcAcStrLn, vbext_pk_Proc) Then
                            For N = (ProcAcStrLn + 1) To (ProcAcStrLn + PrcCnountLine - 1)
                                ThisLine = .Lines(N, 1)
                                If InStr(1, ThisLine, Trim(Fnd), vbTextCompare) > 0 Then
                                    .ReplaceLine N, Replace(ThisLine, Fnd, Rplc, , , vbTextCompare)
                                    Exit For
                                    Exit For
                                    Exit For
                                End If
                            Next N
                        End If
                        Exit Sub
                        Fnd = "": Rplc = ""

                        On Error GoTo 0
                    End If
                End With
            End If
        End With
    Next
End Sub

