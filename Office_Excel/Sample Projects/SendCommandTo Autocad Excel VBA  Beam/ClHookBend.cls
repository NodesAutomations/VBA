VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClHookBend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Coll As Collection

Private Sub Class_Initialize()
 Dim HBdata As ClHookData

  Set HBdata = New ClHookData
  Set Coll = New Collection
 
'     HBdata.Diam = 10
     HBdata.bend = 100
     HBdata.Hook = 120
     HBdata.radius = 30
     Coll.Add HBdata, "10"
       Set HBdata = New ClHookData
'     HBdata.Diam = 12
     HBdata.Hook = 160
     HBdata.bend = 100
     HBdata.radius = 36
     Coll.Add HBdata, "12"
      Set HBdata = New ClHookData
'     HBdata.Diam = 16
     HBdata.Hook = 200
     HBdata.bend = 120
     HBdata.radius = 48
     Coll.Add HBdata, "16"
       Set HBdata = New ClHookData
'     HBdata.Diam = 20
     HBdata.Hook = 240
     HBdata.bend = 140
     HBdata.radius = 60
     Coll.Add HBdata, "20"
        Set HBdata = New ClHookData
'     HBdata.Diam = 25
     HBdata.Hook = 300
     HBdata.bend = 180
     HBdata.radius = 75
     Coll.Add HBdata, "25"
      Set HBdata = New ClHookData
'     HBdata.Diam = 32
     HBdata.Hook = 400
     HBdata.bend = 220
     HBdata.radius = 96
     Coll.Add HBdata, "32"
        Set HBdata = New ClHookData
'     HBdata.Diam = 40
     HBdata.Hook = 480
     HBdata.bend = 260
     HBdata.radius = 120
     Coll.Add HBdata, "40"
       
End Sub

Public Property Get Items(Index As Variant) As ClHookData
'MsgBox Coll.Count
     Set Items = Coll.Item(Index)
End Property

