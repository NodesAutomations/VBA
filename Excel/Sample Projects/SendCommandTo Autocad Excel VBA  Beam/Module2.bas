Attribute VB_Name = "Module2"
Option Explicit
Dim bend As New ClHookBend
Sub Drawbeamttest()

Dim AutocadApp As Object
  Dim Beambody As Object
  Dim Toppart(0 To 7) As Double
  Dim Bottpart(0 To 7) As Double
  Dim Leftpart(0 To 3) As Double
  Dim Rightpart(0 To 3) As Double
  Dim StrTitle As Object
  Dim textpos(2) As Double
  Dim ActDoc As Object
  Dim Blength As Double
  Dim Bwidth As Double
  Dim BHeight As Double
  Dim ColWidthLeft As Double
  Dim ColWidthRight As Double
  Dim Origin(2) As Double
  Dim Topbar As Integer
  Dim Bottbar As Integer
  Dim BottbarSize As Integer
  Dim TotbarSize As Integer
  Dim Cover As Integer
  Dim StirrSpacing As Integer
  Dim StirrSize As Integer
  Dim Ltopreinf(0 To 7)  As Double
  Dim LBotreinf(0 To 7)  As Double
  Dim Rtopreinf(0 To 7)  As Double
  Dim RBotreinf(0 To 7)  As Double
  Dim SteelRef As Object
  Dim Stirrup(0 To 3) As Double, Dimobj As Object
  Dim Pt1(2) As Double, Pt2(2) As Double, Dist As Double, pt3(2) As Double, pt4(2) As Double
  Dim i As Integer, Nstirr As Integer, Stirlenght As Integer, SpacingVAr As Double
  Dim LPt1(2) As Variant, LPt2(2) As Variant, Lpt3(2) As Variant, Lpt4(2) As Variant
  Dim WholeNUM As Integer, Remain As Double, Startdis As Double
  Dim Colour As Object
  '****** Launch Autocad application****
  On Error Resume Next
  Set AutocadApp = GetObject(, "Autocad.application")
  On Error GoTo 0
  
  If AutocadApp Is Nothing Then
  Set AutocadApp = CreateObject("Autocad.application")
  AutocadApp.Visible = True
  End If
 
  Set ActDoc = AutocadApp.ActiveDocument
  
  If ActDoc Is Nothing Then
      Set ActDoc = AutocadApp.Documents.Add
  End If
 ''****Read Beam input****
  Blength = ActiveSheet.Range("b10").value
  BHeight = ActiveSheet.Range("b11").value
  Bwidth = ActiveSheet.Range("b12").value
  ColWidthLeft = ActiveSheet.Range("b13").value
  ColWidthRight = ActiveSheet.Range("b14").value
  Origin(0) = ActiveSheet.Range("D17")
  Origin(1) = ActiveSheet.Range("E17")
  Topbar = ActiveSheet.Range("B19").value
  TotbarSize = ActiveSheet.Range("B20").value
  Bottbar = ActiveSheet.Range("B21").value
  BottbarSize = ActiveSheet.Range("B22").value
  Cover = ActiveSheet.Range("B23").value
  
''****Draw Beam body****
  Leftpart(0) = Origin(0): Leftpart(1) = Origin(1)
  Leftpart(2) = Origin(0): Leftpart(3) = Origin(1) + BHeight * 2
    
     Set Beambody = ActDoc.ModelSpace.AddLightWeightPolyline(Leftpart)
  
  Toppart(0) = Origin(0) + ColWidthLeft: Toppart(1) = Origin(1) + BHeight * 2
  Toppart(2) = Origin(0) + ColWidthLeft: Toppart(3) = Origin(1) + BHeight / 2 + BHeight
  Toppart(4) = Origin(0) + ColWidthLeft + Blength: Toppart(5) = Origin(1) + BHeight / 2 + BHeight
  Toppart(6) = Origin(0) + ColWidthLeft + Blength: Toppart(7) = Origin(1) + BHeight * 2
     Set Beambody = ActDoc.ModelSpace.AddLightWeightPolyline(Toppart)
  
  Bottpart(0) = Origin(0) + ColWidthLeft: Bottpart(1) = Origin(1)
  Bottpart(2) = Origin(0) + ColWidthLeft: Bottpart(3) = Origin(1) + BHeight / 2
  Bottpart(4) = Origin(0) + ColWidthLeft + Blength: Bottpart(5) = Origin(1) + BHeight / 2
  Bottpart(6) = Origin(0) + ColWidthLeft + Blength: Bottpart(7) = Origin(1)
     Set Beambody = ActDoc.ModelSpace.AddLightWeightPolyline(Bottpart)
  
  Rightpart(0) = Origin(0) + Blength + ColWidthLeft + ColWidthRight: Rightpart(1) = Origin(1)
  Rightpart(2) = Origin(0) + Blength + ColWidthLeft + ColWidthRight: Rightpart(3) = Origin(1) + BHeight * 2
  
     Set Beambody = ActDoc.ModelSpace.AddLightWeightPolyline(Rightpart)
          
 ''****Draw Reinforcement****
     'Left top reinforcement
    Ltopreinf(0) = Origin(0) + Cover: Ltopreinf(1) = Origin(1) + BHeight / 2 + BHeight - Cover - bend.Items(CStr(TotbarSize)).bend
    Ltopreinf(2) = Origin(0) + Cover: Ltopreinf(3) = Origin(1) + BHeight / 2 + BHeight - Cover
    Ltopreinf(4) = Origin(0) + Cover + Blength * 2 / 3: Ltopreinf(5) = Origin(1) + BHeight / 2 + BHeight - Cover
    Ltopreinf(6) = Origin(0) + Cover + Blength * 2 / 3 + 30: Ltopreinf(7) = Origin(1) + BHeight / 2 + BHeight - Cover - 30
    
    Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Ltopreinf)

       SteelRef.Color = 1
      SteelRef.ConstantWidth = 5
      
      
     Call FilletVertex(SteelRef, 1, bend.Items(CStr(TotbarSize)).radius, ActDoc)
      
      LPt1(0) = Origin(0) + Cover + Blength * 1 / 3: LPt1(1) = Origin(1) + BHeight / 2 + BHeight - Cover
      LPt2(0) = Origin(0) + Cover + Blength * 1 / 3: LPt2(1) = Origin(1) + BHeight / 2 + BHeight + BHeight / 2
      Lpt3(0) = Origin(0) + Cover + Blength * 1 / 3 + 150: Lpt3(1) = Origin(1) + BHeight / 2 + BHeight + BHeight / 2
      Lpt4(0) = Origin(0) + Cover + Blength * 1 / 3 + 150: Lpt4(1) = Origin(1) + BHeight / 2 + BHeight + BHeight / 2
          
      Call AddLeader(LPt1, LPt2, Lpt3, Lpt4, Topbar, TotbarSize, ActDoc)
      
     
     'Left Bottom reinforcement
    Ltopreinf(0) = Origin(0) + Cover: Ltopreinf(1) = Origin(1) + BHeight / 2 + Cover + bend.Items(CStr(BottbarSize)).bend
    Ltopreinf(2) = Origin(0) + Cover: Ltopreinf(3) = Origin(1) + BHeight / 2 + Cover
    Ltopreinf(4) = Origin(0) + Cover + Blength * 1 / 3: Ltopreinf(5) = Origin(1) + BHeight / 2 + Cover
    Ltopreinf(6) = Origin(0) + Cover + Blength * 1 / 3 + 30: Ltopreinf(7) = Origin(1) + BHeight / 2 + Cover + 30
    
     Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Ltopreinf)

      SteelRef.Color = 1
      SteelRef.ConstantWidth = 5
     
     Call FilletVertex(SteelRef, 1, bend.Items(CStr(BottbarSize)).radius, ActDoc)
     
      LPt1(0) = Origin(0) + Cover + Blength * 1 / 5: LPt1(1) = Origin(1) + BHeight / 2 + Cover
      LPt2(0) = Origin(0) + Cover + Blength * 1 / 5: LPt2(1) = Origin(1)
      Lpt3(0) = Origin(0) + Cover + Blength * 1 / 5 + 150: Lpt3(1) = Origin(1)
      Lpt4(0) = Origin(0) + Cover + Blength * 1 / 5 + 150: Lpt4(1) = Origin(1)
          
      Call AddLeader(LPt1, LPt2, Lpt3, Lpt4, Bottbar, BottbarSize, ActDoc)
     
     'Right Bottom reinforcement
    Ltopreinf(0) = Origin(0) + Cover + Blength * 1 / 3 - 30 - BottbarSize * 40: Ltopreinf(1) = Origin(1) + BHeight / 2 + Cover + 30
    Ltopreinf(2) = Origin(0) + Cover + Blength * 1 / 3 - BottbarSize * 40: Ltopreinf(3) = Origin(1) + BHeight / 2 + Cover
    Ltopreinf(4) = Origin(0) + Blength + ColWidthLeft + ColWidthRight - Cover: Ltopreinf(5) = Origin(1) + BHeight / 2 + Cover
    Ltopreinf(6) = Origin(0) + Blength + ColWidthLeft + ColWidthRight - Cover: Ltopreinf(7) = Origin(1) + BHeight / 2 + Cover + bend.Items(CStr(BottbarSize)).bend
    
    
     Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Ltopreinf)

       SteelRef.Color = 1
      SteelRef.ConstantWidth = 5
      Call FilletVertex(SteelRef, 2, bend.Items(CStr(BottbarSize)).radius, ActDoc)
      
      LPt1(0) = Origin(0) + Cover + Blength * 2 / 3: LPt1(1) = Origin(1) + BHeight / 2 + Cover
      LPt2(0) = Origin(0) + Cover + Blength * 2 / 3: LPt2(1) = Origin(1)
      Lpt3(0) = Origin(0) + Cover + Blength * 2 / 3 + 150: Lpt3(1) = Origin(1)
      Lpt4(0) = Origin(0) + Cover + Blength * 2 / 3 + 150: Lpt4(1) = Origin(1)
          
      Call AddLeader(LPt1, LPt2, Lpt3, Lpt4, Bottbar, BottbarSize, ActDoc)
 
 
    'Right Top reinforcement
    Ltopreinf(0) = Origin(0) + Cover + Blength * 2 / 3 - 30 - TotbarSize * 40: Ltopreinf(1) = Origin(1) + BHeight / 2 + BHeight - Cover - 30
    Ltopreinf(2) = Origin(0) + Cover + Blength * 2 / 3 - TotbarSize * 40: Ltopreinf(3) = Origin(1) + BHeight / 2 + BHeight - Cover
    Ltopreinf(4) = Origin(0) + Blength + ColWidthLeft + ColWidthRight - Cover: Ltopreinf(5) = Origin(1) + BHeight / 2 + BHeight - Cover
    Ltopreinf(6) = Origin(0) + Blength + ColWidthLeft + ColWidthRight - Cover: Ltopreinf(7) = Origin(1) + BHeight / 2 + BHeight - Cover - bend.Items(CStr(TotbarSize)).bend
    
    
     Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Ltopreinf)

       SteelRef.Color = 1
      SteelRef.ConstantWidth = 5
      
      Call FilletVertex(SteelRef, 2, bend.Items(CStr(TotbarSize)).radius, ActDoc)
      
      LPt1(0) = Origin(0) + Cover + Blength * 2 / 3 + 200: LPt1(1) = Origin(1) + BHeight / 2 + BHeight - Cover
      LPt2(0) = Origin(0) + Cover + Blength * 2 / 3 + 200: LPt2(1) = Origin(1) + BHeight / 2 + BHeight + BHeight / 2
      Lpt3(0) = Origin(0) + Cover + Blength * 2 / 3 + 350: Lpt3(1) = Origin(1) + BHeight / 2 + BHeight + BHeight / 2
      Lpt4(0) = Origin(0) + Cover + Blength * 2 / 3 + 350: Lpt4(1) = Origin(1) + BHeight / 2 + BHeight + BHeight / 2
          
      Call AddLeader(LPt1, LPt2, Lpt3, Lpt4, Topbar, TotbarSize, ActDoc)
    
      
 ''*************Add Stirrup ***********
  'Left Support
  
  StirrSpacing = ActiveSheet.Range("B25").value
  StirrSize = ActiveSheet.Range("B26").value
  Stirlenght = ActiveSheet.Range("B27").value
  
  Nstirr = WorksheetFunction.Ceiling((Stirlenght / StirrSpacing), 1)
  WholeNUM = Fix(Stirlenght / StirrSpacing) * StirrSpacing
  Remain = Stirlenght - WholeNUM
  Do Until Startdis > (WholeNUM + Remain)
  
    Stirrup(0) = Origin(0) + ColWidthLeft + Startdis: Stirrup(1) = Origin(1) + BHeight / 2 + Cover
    Stirrup(2) = Origin(0) + ColWidthLeft + Startdis: Stirrup(3) = Origin(1) + BHeight / 2 + BHeight - Cover

 Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Stirrup)

       SteelRef.Color = 1
      SteelRef.ConstantWidth = 5

   If Startdis >= WholeNUM And Remain > 0 Then
     Startdis = Startdis + Remain
  
    Else
     Startdis = Startdis + StirrSpacing
   End If
  Loop
  
  
  Pt1(0) = Origin(0) + ColWidthLeft: Pt1(1) = Origin(1) + BHeight / 4
  Pt2(0) = Origin(0) + ColWidthLeft + Stirlenght: Pt2(1) = Origin(1) + BHeight / 4
  Dist = Distance(Pt1, Pt2)
  
  textpos(0) = Origin(0) + ColWidthLeft + Dist / 2: textpos(1) = Origin(1) + BHeight / 4
 
   Set Dimobj = ActDoc.ModelSpace.AddDimAligned(Pt1, Pt2, textpos)
    Call StirrupDimstyle(Dimobj, Nstirr, StirrSize, StirrSpacing)
 
 'Middle Support
  
  StirrSpacing = ActiveSheet.Range("B29").value
  StirrSize = ActiveSheet.Range("B30").value
  Stirlenght = ActiveSheet.Range("B31").value
  
  Nstirr = WorksheetFunction.Ceiling((Stirlenght / StirrSpacing), 1)
  WholeNUM = Fix(Stirlenght / StirrSpacing) * StirrSpacing
  Remain = Stirlenght - WholeNUM
  Startdis = 0
  
  Do Until Startdis > (WholeNUM + Remain)
  
    Stirrup(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B27").value + Startdis: Stirrup(1) = Origin(1) + BHeight / 2 + Cover
    Stirrup(2) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B27").value + Startdis: Stirrup(3) = Origin(1) + BHeight / 2 + BHeight - Cover

 Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Stirrup)

       SteelRef.Color = 1
      SteelRef.ConstantWidth = 5

  
   If Startdis >= WholeNUM And Remain > 0 Then
     Startdis = Startdis + Remain
  
    Else
     Startdis = Startdis + StirrSpacing
   End If
  Loop
  
  Pt1(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B27").value: Pt1(1) = Origin(1) + BHeight / 4
  Pt2(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B27").value + Stirlenght: Pt2(1) = Origin(1) + BHeight / 4
  Dist = Distance(Pt1, Pt2)
  
  textpos(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B27").value + Dist / 2: textpos(1) = Origin(1) + BHeight / 4
 
   Set Dimobj = ActDoc.ModelSpace.AddDimAligned(Pt1, Pt2, textpos)
   Call StirrupDimstyle(Dimobj, Nstirr, StirrSize, StirrSpacing)

 'Right Support
  
  StirrSpacing = ActiveSheet.Range("B33").value
  StirrSize = ActiveSheet.Range("B34").value
  Stirlenght = ActiveSheet.Range("B35").value
  Startdis = 0
  Nstirr = WorksheetFunction.Ceiling((Stirlenght / StirrSpacing), 1)
  WholeNUM = Fix(Stirlenght / StirrSpacing) * StirrSpacing
  Remain = Stirlenght - WholeNUM
  
  Do Until Startdis > (WholeNUM + Remain)
  
    Stirrup(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B31").value + ActiveSheet.Range("B27").value + Startdis: Stirrup(1) = Origin(1) + BHeight / 2 + Cover
    Stirrup(2) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B31").value + ActiveSheet.Range("B27").value + Startdis: Stirrup(3) = Origin(1) + BHeight / 2 + BHeight - Cover

 Set SteelRef = ActDoc.ModelSpace.AddLightWeightPolyline(Stirrup)

       SteelRef.Color = 1
      SteelRef.ConstantWidth = 5

 If Startdis >= WholeNUM And Remain > 0 Then
     Startdis = Startdis + Remain
  
    Else
     Startdis = Startdis + StirrSpacing
   End If
  
  Loop
  
  Pt1(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B31").value + ActiveSheet.Range("B27").value: Pt1(1) = Origin(1) + BHeight / 4
  Pt2(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B31").value + ActiveSheet.Range("B27").value + Stirlenght: Pt2(1) = Origin(1) + BHeight / 4
  Dist = Distance(Pt1, Pt2)
  
  textpos(0) = Origin(0) + ColWidthLeft + ActiveSheet.Range("B31").value + Dist / 2: textpos(1) = Origin(1) + BHeight / 4
 
   Set Dimobj = ActDoc.ModelSpace.AddDimAligned(Pt1, Pt2, textpos)
   Call StirrupDimstyle(Dimobj, Nstirr, StirrSize, StirrSpacing)
  
  
'ActDoc.Activate

AutocadApp.ZoomExtents
Set Beambody = Nothing
Set AutocadApp = Nothing
Set ActDoc = Nothing
Set SteelRef = Nothing
Set Dimobj = Nothing
End Sub

Function AddLeader(Pt1 As Variant, Pt2 As Variant, pt3 As Variant, pt4 As Variant, Nbar As Integer, Size As Integer, App As Object)
Dim Inspoint(0 To 8) As Double
Dim Myleader As Object

Inspoint(0) = Pt1(0): Inspoint(1) = Pt1(1): Inspoint(2) = Pt1(2)
Inspoint(3) = Pt2(0): Inspoint(4) = Pt2(1): Inspoint(5) = Pt1(2)
Inspoint(6) = pt3(0): Inspoint(7) = pt3(1): Inspoint(8) = Pt1(2)

Set Myleader = App.ModelSpace.AddMLeader(Inspoint, 0)

Myleader.TextString = Nbar & "Y" & Size

Myleader.Color = 3

    With Myleader
        .ArrowheadSize = 40
        .TextHeight = 40
    End With
    
    
End Function


Function Distance(Pt1 As Variant, Pt2 As Variant)
 Dim dx As Double, dy As Variant
   
 dx = Pt2(0) - Pt1(0)
 dy = Pt2(1) - Pt1(1)
 
 Distance = (dx ^ 2 + dy ^ 2) ^ 0.5
 
End Function
Public Function FilletVertex(ByVal LWPline As Object, _
ByVal VertexNumber As Integer, ByVal radius As Double, appdoc As Object)
On Error Resume Next

Dim AngleToVertex As Double
Dim AngleFromVertex As Double
Dim AngleIncluded As Double
Dim PtList As Variant
Dim lastvertex As Integer
Dim PrevVertex As Integer
Dim NextVertex As Integer
Dim Pt1 As Variant
Dim Pt2 As Variant
Dim pt2a As Variant
Dim pt2b As Variant
Dim pt3 As Variant
Dim VertexA(1) As Double
Dim VertexB(1) As Double
Dim AngleC As Double
Dim AngleA As Double
Dim Chamfer As Double
Dim Util As Object
Const PI = 3.14159265358979
Set Util = appdoc.Utility

If Not radius > 0 Then Exit Function

With LWPline
PtList = .Coordinates

lastvertex = (UBound(PtList) - 1) / 2
If VertexNumber > lastvertex Then VertexNumber = 0
NextVertex = VertexNumber + 1
PrevVertex = VertexNumber - 1
If NextVertex > lastvertex Then NextVertex = 0
If PrevVertex < 0 Then PrevVertex = lastvertex

If NextVertex = PrevVertex Then Exit Function

Pt1 = .Coordinate(PrevVertex)
Pt2 = .Coordinate(VertexNumber)
pt3 = .Coordinate(NextVertex)

ReDim Preserve Pt1(2): Pt1(2) = 0
ReDim Preserve Pt2(2): Pt2(2) = 0
ReDim Preserve pt3(2): pt3(2) = 0

AngleToVertex = Util.AngleFromXAxis(Pt2, Pt1)
AngleFromVertex = Util.AngleFromXAxis(Pt2, pt3)

AngleIncluded = (AngleToVertex - AngleFromVertex)
If AngleIncluded > PI Then
AngleIncluded = AngleIncluded - (2 * PI)
ElseIf AngleIncluded < -PI Then
AngleIncluded = AngleIncluded + (2 * PI)
End If

Chamfer = radius * Tan((PI - (Abs(AngleIncluded))) / 2)

pt2b = Util.PolarPoint(Pt2, AngleFromVertex, Chamfer)
VertexB(0) = pt2b(0): VertexB(1) = pt2b(1)
.Coordinate(VertexNumber) = VertexB

pt2a = Util.PolarPoint(Pt2, AngleToVertex, Chamfer)
VertexA(0) = pt2a(0): VertexA(1) = pt2a(1)
.AddVertex VertexNumber, VertexA

.SetBulge VertexNumber, Tan((IIf(AngleIncluded > 0, PI, -PI) - AngleIncluded) / 4#)

End With
End Function
Sub StirrupDimstyle(Dimobj As Object, Nbar As Integer, Size As Integer, spacing As Integer)
 
 With Dimobj
 
 .Color = 7
.ExtensionLineExtend = 30
.Arrowhead1Type = 2
.Arrowhead2Type = 2
.ArrowheadSize = 40
.TextColor = 3
.TextHeight = 40
.UnitsFormat = 2
.PrimaryUnitsPrecision = 0
.TextGap = 15
.LinearScaleFactor = 1
.ExtensionLineOffset = 70
.VerticalTextPosition = 1
.DimLine1Suppress = False
.DimLine2Suppress = False
.ExtLine1Suppress = False
.ExtLine2Suppress = False
.DimTxtDirection = False
.HorizontalTextPosition = 0
.ExtLineFixedLen = 30
.DimensionLineColor = 7
.ExtensionLineColor = 7
.TextOverride = Nbar & " Y " & Size & " @ " & spacing & " C/C"
 End With
End Sub
