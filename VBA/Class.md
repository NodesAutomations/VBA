# Class Modules

### Use Cases

### Code Snippets
```vba
'@Folder("VBAProject")
Option Explicit
 
Private id As Integer
Private CoordinateX As Double
Private CoordinateY As Double
Private CoordinateZ As Double

'node id
Public Property Get uId() As Integer
    uId = id
End Property

Public Property Let uId(value As Integer)
    id = value
End Property

'x coordinate
Public Property Get x() As Double
    x = CoordinateX
End Property

Public Property Let x(value As Double)
    CoordinateX = Round(value, 4)
End Property

'y coordinate
Public Property Get y() As Double
    y = CoordinateY
End Property

Public Property Let y(value As Double)
    CoordinateY = Round(value, 4)
End Property

'z coordinate
Public Property Get z() As Double
    z = CoordinateZ
End Property

Public Property Let z(value As Double)
    CoordinateZ = Round(value, 4)
End Property

Public Sub Display()
    Debug.Print , uId, x, y, z
End Sub

Public Function ToString() As String
    ToString = CStr(uId) & "|" + CStr(x) & "," & CStr(y) & "," & CStr(z)
End Function
```



### Reference
- [Macro Mastery; Class Module](https://excelmacromastery.com/vba-class-modules/)
