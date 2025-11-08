## Embedded Projects

An AutoCAD® VBA project is a collection of code modules, class modules, and forms that work together to perform a given function. Projects can be stored within an AutoCAD drawing, or as a separate file.

**Embedded projects are stored within an AutoCAD drawing.** These projects are automatically loaded whenever the drawing in which they are contained is opened in AutoCAD, making the distribution of projects very convenient. Embedded projects are limited and not able to open or close AutoCAD drawings because they function only within the document where they reside. Users of embedded projects are no longer required to find and load project files before they run a program. A time log that is triggered when the drawing is opened is an example of a project embedded in a drawing. With this macro users can log in and record the length of time they worked on the drawing. The user does not have to remember to load the project before opening the drawing; it simply is done automatically.

Global projects are stored in separate files and are more versatile because they can work in, open, and close any AutoCAD drawing, but are not automatically loaded when a drawing is opened. Users must know which project file contains the macro they need and then load that project file before they can run the macro. However, global projects are easier to share with other users, and they make excellent libraries for common macros. An example of a project you may store in a project file is a macro that collects a bill of materials from many drawings. This macro can be run by an administrator at the end of a work cycle and can collect information from many drawings.

At any given time, users can have both embedded and global projects loaded into their AutoCAD session

> When you embed a project you place a copy of the project in the drawing database. The project is then loaded or unloaded whenever the drawing containing it is opened or closed. A drawing can contain only one embedded project at a time. If a drawing already contains an embedded project you must extract it before a different project can be embedded into the drawing.
> 

### To Embed a Project in an AutoCAD Drawing

Embedding a VBA project file into a drawing allows you to ensure a set of macros are available when the drawing is opened each time.
1. On the ribbon, click Manage tab  Applications panel (expanded)  VBA Manager.
2. In the VBA Manager, select the project you want to embed.
3. Click Embed.

### To Extract a Project from an AutoCAD Drawing

An embedded VBA project can be removed and saved to a DVB file that can be loaded into other drawings.

1. On the ribbon, click Manage tab> Applications panel (expanded) > VBA Manager.
2. In the VBA Manager, click Extract.
The Extract button is only enabled when a drawing contains an embedded VBA project file.
3. In the AutoCAD message box, click Yes to export the VBA project to a DVB file and remove the embedded project file. Click No to just remove the embedded project file.
4. If you clicked Yes, specify a name and location for the DVB file. Click Save.

## Global Projects
