Attribute VB_Name = "Module1"
Option Explicit

Sub DrawRectange()
  Dim AutocadApp As Object
  Dim SectionCoord(0 To 9) As Double
  Dim Topbar As Integer
  Dim BottomBar As Integer
  Dim Cover As Integer
  Dim Rectang As Object
  Dim ActDoc As Object
  Dim InsertP(2) As Double
  Dim CirObj As Object
  Dim i As Long
  Dim spacing As Double
  Dim Topspacing As Double
  Dim TopbarSize As Integer
  Dim BotbarSize As Integer
  Dim MidBar As Integer
  Dim midSize As Integer
  Dim Offrect As Variant
  Dim Stirrup As Object
  Dim FilledCir As Object
  Dim MyArr(0) As Object
  '****** Launch Autocad application****
  On Error Resume Next
  Set AutocadApp = GetObject(, "Autocad.application")
  On Error GoTo 0
  
  If AutocadApp Is Nothing Then
  Set AutocadApp = CreateObject("Autocad.application")
  AutocadApp.Visible = True
  End If
 
 ''****Read Input****
   SectionCoord(0) = 0: SectionCoord(1) = 0
   SectionCoord(2) = ActiveSheet.Range("f5").value: SectionCoord(3) = 0
   SectionCoord(4) = ActiveSheet.Range("f5").value: SectionCoord(5) = ActiveSheet.Range("f6").value
   SectionCoord(6) = 0: SectionCoord(7) = ActiveSheet.Range("f6").value
   SectionCoord(8) = 0: SectionCoord(9) = 0
  
  Topbar = ActiveSheet.Range("f8").value
  BottomBar = ActiveSheet.Range("f10").value
  Cover = ActiveSheet.Range("f14").value
  BotbarSize = ActiveSheet.Range("f11").value
  TopbarSize = ActiveSheet.Range("f9").value
  MidBar = ActiveSheet.Range("f12").value
  midSize = ActiveSheet.Range("f13").value

 ''****Draw rectangle****
  Set ActDoc = AutocadApp.ActiveDocument
  
  If ActDoc Is Nothing Then
      Set ActDoc = AutocadApp.Documents.Add
  End If
  
  Set Rectang = ActDoc.ModelSpace.AddLightWeightPolyline(SectionCoord)
        
   Offrect = Rectang.Offset(-Cover)
     
   Set Stirrup = Offrect(0)
     
   Stirrup.ConstantWidth = 5
   
    spacing = ((ActiveSheet.Range("f5") - 2 * (Cover + BotbarSize / 2 + 5))) / (BottomBar - 1)
   For i = 1 To BottomBar
        If i = 1 Then
           InsertP(0) = Cover + BotbarSize / 2 + 5: InsertP(1) = Cover + BotbarSize / 2 + 5
          
          Else
            InsertP(0) = Cover + BotbarSize / 2 + 5 + spacing * (i - 1): InsertP(1) = Cover + BotbarSize / 2 + 5
        End If
      Set CirObj = ActDoc.ModelSpace.AddCircle(InsertP, BotbarSize / 2)
        CirObj.Color = acred
        
       Set FilledCir = ActDoc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "Solid", True)
       Set MyArr(0) = CirObj
         With FilledCir
            .AppendOuterLoop MyArr
            .Evaluate
            .Color = acred
            .Update
         End With
   Next i
        
    Topspacing = ((ActiveSheet.Range("f5") - 2 * (Cover + TopbarSize / 2 + 5))) / (Topbar - 1)
   For i = 1 To Topbar
        If i = 1 Then
           InsertP(0) = Cover + TopbarSize / 2 + 5: InsertP(1) = ActiveSheet.Range("f6") - Cover - TopbarSize / 2 - 5
          
          Else
            InsertP(0) = Cover + TopbarSize / 2 + 5 + Topspacing * (i - 1): InsertP(1) = ActiveSheet.Range("f6") - Cover - TopbarSize / 2 - 5
        End If
      Set CirObj = ActDoc.ModelSpace.AddCircle(InsertP, TopbarSize / 2)
       CirObj.Color = acred
       Set FilledCir = ActDoc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "Solid", True)
      Set MyArr(0) = CirObj
         With FilledCir
            .AppendOuterLoop (MyArr)
            .Evaluate
            .Color = acred
            .Update
            
         End With
       
       
   Next i
        
     If MidBar <> 0 And MidBar = 2 Then
           InsertP(0) = Cover + midSize / 2 + 5: InsertP(1) = ActiveSheet.Range("f6") / 2
           Set CirObj = ActDoc.ModelSpace.AddCircle(InsertP, midSize / 2)
            CirObj.Color = acred
       
       Set FilledCir = ActDoc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "Solid", True)
         Set MyArr(0) = CirObj
         With FilledCir
            .AppendOuterLoop (MyArr)
            .Evaluate
            .Color = acred
            .Update
         End With
           
           InsertP(0) = Cover + midSize / 2 + (ActiveSheet.Range("f5") - Cover * 2 - midSize - 5): InsertP(1) = ActiveSheet.Range("f6") / 2
       
           Set CirObj = ActDoc.ModelSpace.AddCircle(InsertP, midSize / 2)
           CirObj.Color = acred
       Set FilledCir = ActDoc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "Solid", True)
         Set MyArr(0) = CirObj
         With FilledCir
            .AppendOuterLoop (MyArr)
            .Evaluate
            .Color = acred
            .Update
         End With
     
     End If
        
    AutocadApp.ZoomExtents

Set AutocadApp = Nothing
Set ActDoc = Nothing
Set Rectang = Nothing
Set CirObj = Nothing
Set MyArr(0) = Nothing
Set FilledCir = Nothing
End Sub
