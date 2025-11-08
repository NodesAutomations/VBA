## Overview
- Xdata used to assign extended data to any entity in autocad
- For example, If we can assign Pipe Properties to Line Object in AutoCAD
- This is really good method to store additional info on any AutoCAD Object

## How to use XData Manually on autocad
- Use `XDATA` Command to assign extended data to selected object, you also have to provide Application Name which is used to get data later
- Use `XDLIST` Command to view XData Assign with specific object, you really need to know name of application to view data
- XData is generally more suited for attaching to objects.  XData is indexed by registered applications (RegApps), so you can assign data for one "application" without interfering with others.  There is a limit to the total XData you can assign to any given object but its quite big; if you overrun the limit its a good indication there's a better way to do what you are trying.
- [XDATA (Express Tool) | AutoCAD 2021 | Autodesk Knowledge Network](https://knowledge.autodesk.com/support/autocad/learn-explore/caas/CloudHelp/cloudhelp/2021/ENU/AutoCAD-Core/files/GUID-F0299B36-232F-446E-9F81-98F300B36991-htm.html)
- [xdata vs xrecords? (theswamp.org)](https://www.theswamp.org/index.php?topic=38961.0)
