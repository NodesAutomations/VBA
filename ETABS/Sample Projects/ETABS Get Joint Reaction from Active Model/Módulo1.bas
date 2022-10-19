Attribute VB_Name = "Módulo1"
    Option Explicit

    'Get a reference to cSapModel to access all OAPI classes and functions
    Dim SapModel As ETABSv1.cSapModel
    
    'Late bind to ETABS.exe, create an instance of ETABSObject, and get a reference to cOAPI interface
    Dim ETABSObject As ETABSv1.cOAPI
    Dim myHelper As ETABSv1.cHelper

       Dim ret As Integer
       Dim NumberResults As Long
       Dim Obj() As String
       Dim Elm() As String
       Dim LoadCase() As String
       Dim StepType() As String
       Dim StepNum() As Double
       Dim F1() As Double
       Dim F2() As Double
       Dim F3() As Double
       Dim M1() As Double
       Dim M2() As Double
       Dim M3() As Double
       Dim prev_data As Integer
       Dim NumberNames As Long
       Dim myFile As String
       Dim MyName() As String
       Dim i As Integer
       
       Dim stt, eend As Integer
             

Public Sub ETABS_Attaching()

'Create the ETABS object
Set ETABSObject = GetObject(, "CSI.ETABS.API.ETABSObject")

'Since the program is already started, there is no need to call ETABSObject.ApplicationStart
Set SapModel = ETABSObject.SapModel

'Check attachment
eend = InStrRev(SapModel.GetModelFilename(), ".")
stt = InStrRev(SapModel.GetModelFilename(), "\")
MsgBox ("Attached to model: " & Mid(SapModel.GetModelFilename(), stt + 1, eend - stt - 1))

   'run analysis
       ret = SapModel.Analyze.RunAnalysis
     
    'Set present units to kN-m
        ret = SapModel.SetPresentUnits(6)
        
    'Get Load Cases
        ret = SapModel.LoadCases.GetNameList(NumberNames, MyName)
        For i = 0 To NumberNames - 1
            Sheets(1).ListBox1.AddItem MyName(i)
        Next i

   End Sub

Public Sub Get_Reactions()

    'Clear previous data
    If Cells(2, 1) <> "" Then
        prev_data = Cells(1, 1).End(xlDown).Row
        Range(Cells(2, 1), Cells(prev_data, 8)).ClearContents
    End If
        
    'Get Point Object names
    ret = SapModel.PointObj.GetNameListOnStory("Base", NumberNames, MyName)

   'set case selected for output
     ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(Sheets(1).ListBox1.List(Sheets(1).ListBox1.ListIndex))

   'get point displacements
    For i = 0 To NumberNames - 1
       ret = SapModel.Results.JointReact(MyName(i), ETABSv1.eItemType_Objects, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)
        Cells(i + 2, 1) = MyName(i)
        Cells(i + 2, 2) = Sheets(1).ListBox1.List(Sheets(1).ListBox1.ListIndex)
        Cells(i + 2, 3) = F1(0)
        Cells(i + 2, 4) = F2(0)
        Cells(i + 2, 5) = F3(0)
        Cells(i + 2, 6) = M1(0)
        Cells(i + 2, 7) = M2(0)
        Cells(i + 2, 8) = M3(0)
    Next i

End Sub

Public Sub Close_ETABS()

    'Clear listbox
        Sheets(1).ListBox1.Clear

    'Clear previous data
        If Cells(2, 1) <> "" Then
            prev_data = Cells(1, 1).End(xlDown).Row
            Range(Cells(2, 1), Cells(prev_data, 8)).ClearContents
        End If

   'close ETABS
       ETABSObject.ApplicationExit (False)

   'clean up variables
       Set SapModel = Nothing
       Set ETABSObject = Nothing

End Sub

