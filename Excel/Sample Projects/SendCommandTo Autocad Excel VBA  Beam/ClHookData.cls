VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClHookData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private dHook As Long
Private dBend As Long
Private dRadius As Long
Private dDiam As Long

 Public Property Get Hook() As Long
         Hook = dHook
 End Property
Public Property Let Hook(value As Long)
         dHook = value
End Property
Public Property Get bend() As Long
        bend = dBend
 End Property
Public Property Let bend(value As Long)
         dBend = value
End Property

Public Property Get radius() As Long
        radius = dRadius
End Property
Public Property Let radius(value As Long)
         dRadius = value
End Property

