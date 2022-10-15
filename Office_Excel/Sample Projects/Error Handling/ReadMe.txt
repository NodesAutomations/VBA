ExcelMacroMastery.com Error Handling

INSTALLATION
============

You can install the Error Handling in two simple ways:

For both methods
1. Open the workbook where you want to add the error handling.
2. Go to the Visual Basic editor(Alt + F11).

Method 1:
1. In the Project Window(normally on the left), right-click on the workbook where you want to add the error handling.
2. Select Import.
3. Select the ErrorHandling.bas file and press Ok.

Method 2: 
1. Drag the ErrorHandling.bas file from Windows Explorer into the appropriate workbook in the Project Window.

Note: When the errorhandling module is in a workbook you also can copy to another workbook by dragging it in the Project explorer window.


HOW TO USE THE ERROR HANDLER
============================

Using the error handler in your code is simple:

1. Place DisplayError in the topmost sub at the bottom. Replace the third parameter
  with the name of the sub:

DisplayError Err.source, Err.Description, "Module1.Topmost", Erl

2. Place RaiseError in all the other subs at the bottom of each. Replace the third parameter
  with the name of the sub:

RaiseError Err.Number, Err.source, "Module1.Level1", Err.Description, Erl


3. The error handling in each sub should look like this:

Sub subTopmost()

  On Error Goto eh

  The main code of the sub here!!!!!

done:
    Exit Sub
eh:
    DisplayError Err.Source, Err.Description, "Module1.Topmost", Erl
End Sub


Sub subLevel2()

  On Error Goto eh

  The main code of the sub here!!!!!

done:
    Exit Sub
eh:
    RaiseError Err.Number, Err.source, "Module1.Level2", Err.Description, Erl
End Sub
