## Overview
- Xdata used to assign extended data to any entity in autocad
- For example, If we can assign Pipe Properties to Line Object in AutoCAD
- This is really good method to store additional info on any AutoCAD Object

## How to use XData Manually on autocad
- Use `XDATA` Command to assign extended data to selected object, you also have to provide Application Name which is used to get data later
- Use `XDLIST` Command to view XData Assign with specific object, you really need to know name of application to view data
- Refer : [AutoCAD XDATA](https://knowledge.autodesk.com/support/autocad/learn-explore/caas/CloudHelp/cloudhelp/2021/ENU/AutoCAD-Core/files/GUID-F0299B36-232F-446E-9F81-98F300B36991-htm.html)

### Title

Hey this is just a sample to see if this works with github or not

### Title2

- this looks really fun
- and this works

### Title

Hey this is just a sample to see if this works with github or not

### Title2

- this looks really fun
- and this works

```vbnet
Sub Test()
'This is just a sample code
End Sub
```
| Application name | An ASCII string up to 255 bytes long (group code 1000). |
| --- | --- |
| Layer | A layer name (group code 1003). |
| Hand | An object handle (group code 1005). |
| 3Real | 3 real numbers (group code 1010). |
| Pos | A 3D World space position (group code 1011). |
| Disp | A 3D World space displacement (group code 1012). |
| Dir | A 3D World space direction (group code 1013). |
| Real | A real number (group code 1040). |
| Dist | A distance (group code 1041). |
| Scale | A scale factor (group code 1042). |
| Int | A 16-bit integer (group code 1070). |
| Long | A 32-bit signed long integer (group code 1071). |
