Attribute VB_Name = "mACADText"
Option Explicit

Sub AddText()

    '----------------------------------------------------------------------------------------------------------------
    'This macro adds text in AutoCAD using data - insertion point, text height and text message - from Excel.
    'Moreover, it provides some optional parameters, so that the user can adjust the alignment, the rotation angle,
    'the width factor and the color of the text to be inserted (using the Red, Green and Blue parameters).
    'The code uses late binding, so no reference to external AutoCAD (type) library is required.
    'It goes without saying that AutoCAD must be installed at your computer before running this code.
    
    'Written By:    Christos Samaras
    'Date:          07/03/2014
    'Last Update:   16/03/2015
    'E-mail:        xristos.samaras@gmail.com
    'Site:          http://www.myengineeringworld.net
    '----------------------------------------------------------------------------------------------------------------
        
    'Declaring the necessary variables.
    Dim acadApp                 As Object
    Dim acadDoc                 As Object
    Dim acadText                As Object
    Dim acadColor               As Object
    Dim LastRow                 As Long
    Dim i                       As Long
    Dim InsertionPoint(0 To 2)  As Double
    Dim ZeroPoint(0 To 2)       As Double
    
    'Activate the coordinates sheet and find the last row.
    With Sheets("Coordinates")
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
        
    'Check if there are coordinates for at least one text message.
    If LastRow < 2 Then
        MsgBox "There are no coordinates for the insertion point!", vbCritical, "Insertion Point Error"
        Exit Sub
    End If
    
    'Check if AutoCAD application is open. If is not opened create a new instance and make it visible.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("AutoCAD.Application")
        acadApp.Visible = True
    End If
    
    'Check (again) if there is an AutoCAD object.
    If acadApp Is Nothing Then
        MsgBox "Sorry, it was impossible to start AutoCAD!", vbCritical, "AutoCAD Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    'If there is no active drawing create a new one.
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
    End If
    On Error GoTo 0

    'Check if the active space is paper space and change it to model space.
    If acadDoc.ActiveSpace = 0 Then '0 = acPaperSpace in early binding
        acadDoc.ActiveSpace = 1     '1 = acModelSpace in early binding
    End If
             
    'Set the AcCmColor object (here acadColor) which represents colors.
    Set acadColor = acadApp.GetInterfaceObject("AutoCAD.AcCmColor." & Left(acadApp.Version, 2))
                
    'The point at the beginning of 3 axes.
    ZeroPoint(0) = 0
    ZeroPoint(1) = 0
    ZeroPoint(2) = 0
    
    'Loop through all the rows and add the corresponding text in AutoCAD.
    With Sheets("Coordinates")
        For i = 2 To LastRow
            'If the height and the message are not empty, add the text.
            If IsEmpty(.Range("D" & i)) = False And IsEmpty(.Range("E" & i)) = False Then
                'Set the insertion point.
                InsertionPoint(0) = .Range("A" & i).Value
                InsertionPoint(1) = .Range("B" & i).Value
                InsertionPoint(2) = .Range("C" & i).Value
                'Add the text.
                Set acadText = acadDoc.ModelSpace.AddText(.Range("E" & i), InsertionPoint, .Range("D" & i))
                'Align the text based on alignment selection. The default is left.
                If IsEmpty(.Range("F" & i)) = False Then
                    If UCase(.Range("F" & i)) = "CENTER" Then
                        acadText.Alignment = 1
                        acadText.Move ZeroPoint, InsertionPoint
                    End If
                    If UCase(.Range("F" & i)) = "RIGHT" Then
                        acadText.Alignment = 2
                        acadText.Move ZeroPoint, InsertionPoint
                    End If
                End If
                'If the rotation angle is not empty, rotate the text.
                '0.0174532925 is used to convert from degrees to radians.
                If IsEmpty(.Range("G" & i)) = False Then acadText.Rotate InsertionPoint, .Range("G" & i).Value * 0.0174532925
                'If the width factor is not empty, apply the width factor.
                If IsEmpty(.Range("H" & i)) = False Then acadText.ScaleFactor = .Range("H" & i)
                'Use the sheet data to change the color on each line of text (if the corresponding cells are not empty).
                If IsEmpty(.Range("I" & i)) = False And IsEmpty(.Range("J" & i)) = False And IsEmpty(.Range("K" & i)) = False Then
                    Call acadColor.SetRGB(.Range("I" & i), .Range("J" & i), .Range("K" & i))
                    'Change the color of text.
                    acadText.TrueColor = acadColor
                End If
            End If
        Next i
    End With
    
    'Zoom in to the drawing area.
    acadApp.ZoomExtents

    'Release the objects.
    Set acadText = Nothing
    Set acadDoc = Nothing
    Set acadApp = Nothing
    
    'Inform the user about the process.
    MsgBox "The text was successfully added in AutoCAD!", vbInformation, "Finished"

End Sub

Sub ClearAll()
    
    Dim LastRow As Long
    
    'Find the last row and clear all the input data..
    With Sheets("Coordinates")
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If LastRow > 1 Then
            .Range("A2:K" & LastRow).ClearContents
        End If
        .Range("A2").Select
    End With
    
End Sub
