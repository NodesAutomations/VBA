### Find Last USed Row Column
```vba
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.range("A" & ws.Rows.Count).End(xlUp).Row
```
### Find Current Region based on specific cell
```vba
Sheet1.Range("E8").CurrentRegion.Address
'Print
'$D$8:$E$10
```
![image](https://user-images.githubusercontent.com/60865708/126315441-33bb5d22-5478-4337-b892-d6561fad3103.png)

### Get Range Data in Arra
```vba
Dim arr as Variant
arr=Sheet1.Range("E8").CurrentRegion.Value
```
### Store Range In Arrray For Faster Processing
```vba
dim rg as Range
set rg=Sheet1.Range("E8").CurrentRegion

'Reading Data From Range
dim arr as Variant
arr=rg.Value

'Writing Data from Range
Sheet1.Range("I8:M18").Value=arr
```
### Name Range
```visual-basic
Sub NameRange_Add()

Dim cell As Range
Dim rng As Range
Dim RangeName As String
Dim CellName As String

'Single Cell Reference (Workbook Scope)
  RangeName = "Price"
  CellName = "D7"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell

'Single Cell Reference (Worksheet Scope)
  RangeName = "Year"
  CellName = "A2"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  Worksheets("Sheet1").Names.Add Name:=RangeName, RefersTo:=cell

'Range of Cells Reference (Workbook Scope)
  RangeName = "myData"
  CellName = "F9:J18"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell

'Secret Named Range (doesn't show up in Name Manager)
  RangeName = "Username"
  CellName = "L45"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell, Visible:=False

End Sub
```

```visual-basic
Sub Update()

Dim cell As Range
Dim CellName As String
Dim RangeName As String
Dim RangeValue As String
Dim SheetName As String

Dim FileNum As Integer
Dim DataLine As String

FileNum = FreeFile()
Open "C:\Users\Ryzen2600x\TestRepo\AddNamedRange\AddNamedRange\MissingVar.txt" For Input As #FileNum
Dim str() As String

While Not EOF(FileNum)
    Line Input #FileNum, DataLine ' read in data 1 line at a time
    str = Split(DataLine, ",")
    'MsgBox (str(2))
        
    SheetName = str(0)
    CellName = str(1)
    RangeName = str(2)
    RangeValue = str(3)

 Set cell = Worksheets(SheetName).Range(CellName)
 If Not RangeName = "" Then
  ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=cell
  Range(RangeName) = RangeValue
  Else
  cell.Value = RangeValue
 End If
Wend
```
## Select A Range With An Input Box
Clipped from: [https://www.thespreadsheetguru.com/blog/vba-to-select-range-with-inputbox](https://www.thespreadsheetguru.com/blog/vba-to-select-range-with-inputbox)

I’ve run into a few times where I felt the user experience would be more streamlined if I gave them the option to bring in data or properties from their spreadsheet. You have a couple of options for doing this, as you can either pull this information from:

- the user’s current selection
- a predetermined cell range or object name
- ask the user to type a cell address into an inputbox

I would argue the last bullet listed above is most likely going to be the most straightforward for your users as they have the most freedom to state any range address they want at that point in time. But is it worth giving the user this much freedom? Is it worth spending the countless hours coding to prevent any incorrect input a user might enter into your inputbox?

This is where the beauty of the built-in VBA InputBox object will save you time and effort. Let’s look at how we can use the InputBox to easily prompt the user to select a cell range so we can store that range location to a variable.

## **Using The InputBox Object**

I won’t go into all the detail of what the InputBox Object can do as you can read all the attributes via Microsoft’s documentation **[here](https://docs.microsoft.com/en-us/office/vba/api/excel.application.inputbox). However, I will note that the InputBox** has some very handy input restrictions that we can use to easily account for all the error handling we need to confirm our users are properly inputting a valid Excel range. This attribute is called the **Type**.

### **InputBox “Type” Attribute**

ValueMeaning

[Untitled](https://www.notion.so/5c0028d01b3f44aeb7580027e96dad3a)

### **InputBox Object Attributes**

Below are all the attributes you may modify while using the InputBox object. The main ones we will be using are the **Prompt**, **Title**, and **Type** attributes.

**Application.InputBox**(*Prompt*, *Title*, *Default*, *Left*, *Top*, *HelpFile*, *HelpContextID*, *Type*)

## **VBA Code Examples**

Effectively what the following VBA code examples are going to carry out is prompt the user to either enter or select a cell range with their cursor. This **InputBox** also has the ability to reference different worksheets within the same workbook file. This flexibility really allows your user to have the best experience while providing your VBA macro a variable cell range to work with.

### **Grab A Cell Range**

In this VBA code example, the macro’s goal will be to retrieve a **Custom Number Format** rule from the user and apply it to the user’s current cell selection. The **InputBox** will be used to gather a single cell input from the user and store that cell and all it’s properties to a variable. This way, the user does not need to type out the number format rule themselves, they can simple point to a cell that already has the rule applied. This technique is also extremely useful to getting color inputs from your users.

```
Sub NumberFormatFromCell()
 'PURPOSE: Obtain A Number Format Rule From A Cell User's Determines
 
 Dim rng As Range
 Dim FormatRuleInput As String
 
 'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
   Set rng = Application.InputBox( _
    Title:="Number Format Rule From Cell", _
    Prompt:="Select a cell to pull in your number format rule", _
    Type:=8)
  On Error GoTo 0
 
 'Test to ensure User Did not cancel
  If rng Is Nothing Then Exit Sub
  
 'Set Variable to first cell in user's input (ensuring only 1 cell)
  Set rng = rng.Cells(1, 1)
 
 'Store Number Format Rule
  FormatRuleInput = rng.NumberFormat
 
 'Apply NumberFormat To User Selection
  If TypeName(Selection) = "Range" Then
   Selection.NumberFormat = FormatRuleInput
  Else
   MsgBox "Please select a range of cells before running this macro!"
  End If
 
 End Sub

```

### **Grab A Cell Range With A Default**

The below VBA macro code shows you how to display a default cell range value when the InputBox first launches. In this example, the default range will be the users current cell selection.

```
Sub HighlightCells()
 'PURPOSE: Apply A Yellow Fill Color To A Cell Range
 
 Dim rng As Range
 Dim DefaultRange As Range
 Dim FormatRuleInput As String
 
 'Determine a default range based on user's Selection
  If TypeName(Selection) = "Range" Then
   Set DefaultRange = Selection
  Else
   Set DefaultRange = ActiveCell
  End If
 
 'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
   Set rng = Application.InputBox( _
    Title:="Highlight Cells Yellow", _
    Prompt:="Select a cell range to highlight yellow", _
    Default:=DefaultRange.Address, _
    Type:=8)
  On Error GoTo 0
 
 'Test to ensure User Did not cancel
  If rng Is Nothing Then Exit Sub
  
 'Highlight Cell Range
  rng.Interior.Color = vbYellow
 
 End Sub

```

### **Grab A Cell Range Using A Userform**

If you want to use the **InputBox** technique within a userform, I highly recommend hiding your userform before prompting the user with the **InputBox**. The following code is an example of how you might accomplish this.

```
Sub NumberFormatFromCell()
 'PURPOSE: Obtain A Number Format Rule From A Cell User's Determines

 Dim rng As Range
 Dim FormatRuleInput As String

 'Temporarily Hide Userform
  Me.Hide

 'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
   Set rng = Application.InputBox( _
    Title:="Number Format Rule From Cell", _
    Prompt:="Select a cell to pull in your number format rule", _
    Type:=8)
  On Error GoTo 0

 'Test to ensure User Did not cancel
  If rng Is Nothing Then
   Me.Show 'unhide userform
   Exit Sub
  End If

 'Set Variable to first cell in user's input (ensuring only 1 cell)
  Set rng = rng.Cells(1, 1)

 'Store Number Format Rule
  FormatRuleInput = rng.NumberFormat

 'Apply NumberFormat To User Selection
  If TypeName(Selection) = "Range" Then
   Selection.NumberFormat = FormatRuleInput
  Else
   MsgBox "Please select a range of cells before running this macro!"
  End If

 'Unhide Userform
  Me.Show

 End Sub

```

## **How Do I Modify This To Fit My Specific Needs?**

Chances are this post did not give you the exact answer you were looking for. We all have different situations and it's impossible to account for every particular need one might have. That's why I want to share with you: **[My Guide to Getting the Solution to your Problems FAST!](http://www.thespreadsheetguru.com/gethelp)** **In this article, I explain the best strategies I have come up with over the years to getting quick answers to complex problems in Excel, PowerPoint, VBA,** *you name it*!

I highly recommend that you check **[this guide](http://www.thespreadsheetguru.com/gethelp)** **out before asking me or anyone else in the comments section to solve your specific problem. I can guarantee 9 times out of 10, one of my strategies will get you the answer(s) you are needing faster than it will take me to get back to you with a possible solution. I try my best to help everyone out, but sometimes I don't have time to fit everyone's questions in (there never seem to be quite enough hours in the day!).**

I wish you the best of luck and I hope this tutorial gets you heading in the right direction!

Chris
Founder of TheSpreadsheetGuru



