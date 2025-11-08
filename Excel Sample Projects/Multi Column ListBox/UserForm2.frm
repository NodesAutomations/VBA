VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CommandButton1_Click()
Dim curIndex As Long
Dim othIndex As Long

Dim curValue As String
Dim othValue As String

Dim curValueB As String
Dim othValueB As String

With Me.ListBox1
If .ListIndex = 0 Then Exit Sub
    curIndex = .ListIndex
    othIndex = curIndex - 1
    
    curValue = .Column(0, curIndex)
    othValue = .Column(0, othIndex)
    
    curValueB = .Column(1, curIndex)
    othValueB = .Column(1, othIndex)
    
    .Column(0, curIndex) = othValue
    .Column(0, othIndex) = curValue
    
    .Column(1, curIndex) = othValueB
    .Column(1, othIndex) = curValueB
    
    .Selected(othIndex) = True
    

End With
End Sub

Private Sub CommandButton2_Click()
Dim curIndex As Long
Dim othIndex As Long

Dim curValue As String
Dim othValue As String

Dim curValueB As String
Dim othValueB As String

With Me.ListBox1
If .ListIndex = .ListCount - 1 Then Exit Sub
    curIndex = .ListIndex
    othIndex = curIndex + 1
    
    curValue = .Column(0, curIndex)
    othValue = .Column(0, othIndex)
    
    curValueB = .Column(1, curIndex)
    othValueB = .Column(1, othIndex)
    
    .Column(0, curIndex) = othValue
    .Column(0, othIndex) = curValue
    
    .Column(1, curIndex) = othValueB
    .Column(1, othIndex) = curValueB
    
    .Selected(othIndex) = True
    

End With
End Sub

Private Sub UserForm_Activate()
Dim x As Long
x = 2

Do While Cells(x, 1) <> ""

Me.ListBox1.AddItem Cells(x, 1)
Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(x, 2)
x = x + 1

Loop
End Sub
