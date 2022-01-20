INSTALLATION

Installing RefTreeAnalyser requires these simple steps:

1. Open Downloaded Zip file
Open the zip file which you have downloaded from my website.

2. Copy files
Copy all files from the zip file to any folder you like on your system.

3. Open add-in file RefTreeAnalyserXL.xlam (Users with Excel 2003 or older are advised to open the file called RefTreeAnalyser.xla)
Simply double-click on the excel file you just copied from the zip file.

**********************
*Blocked file problem*
**********************

If Excel fails to open the file, this is due to a recent update which blocks files downloaded from the internet without displaying any warning messages. To resolve this:
- Right-click the file and select Properties
- Click the Unblock button or check-box near the bottom of the dialog.

4. Enable macro's:
You can either click Enable, or "Trust All From Publisher". The latter will ensure any future add-ins you download from my website will have their macro's enabled by default.

5. Install as Add-in

After enabling macro's, RefTreeAnalyser will ask you whether or not you wish to install it as an add-in.
Click yes to have the add-in available every time you start Excel.

If nothing happens when you open RefTreeAnalyser, try this to install it:

- Close all Excel windows 
- Confirm in Task manager that no remaining Excel.exe processes are running, if there are, kill them
- Open Excel
- Click File, Options
- Click the Add-ins tab
- Click the Go... button close to the bottom
- Click the Browse button
- Find RefTreeAnalyserXL.xlam and click Open

WHAT'S NEW

Build 172: fixed bug in off-sheet references option

Build 171: Internal build

Build 170: Improved sorting of chart series nodes on Objects dialog

Build 169: Worked around an Excel limitation to fix an issue with the Analyze off-sheet references tool

Build 168: Worked around a bug which Microsoft introduced

Build 167: Fixed a bug which I introduced in build 166 :-(

Build 166: Two updates: 1. The link back to the ToC is now independent of the name of the file 2. If you press Precedents when a chart is selected, you get all references pertaining to that chart

Build 165: Added Tables to the table of content option

Build 164: Added Form and ActiveX controls to the table of content option

Build 163: Ensured obscure reference containing both a table total row and a cell outside the table works

Build 162: Fix for 64 bit Excel

Build 161: Workaround for rare error regarding SAP add-in.

Build 160: Small bugfix to ensure backward compatibility with Excel 2013 and older.

Build 159: Improved display of object references

Build 158: The tool now recognizes external references in charts and no longer ignores them

Build 157: Fixed bug in Search Objects feature: chart titles with formulas are now properly listed with their charts

Build 156: Small fix for 64 bit Excel

Build 155: Added two new options: 1. Generate a Table of Contents and 2. Added a feedback button which takes you to a small survey so you can tell me what you like and what you do not like.

Build 154: Fixed a bug relating to 64 bit Office. This is a recommended update if you are using 64 Office, which nowadays is the default version installed with a Microsoft 365 license

Build 153: Internal version

Build 152: Improved performance and added an option to prevent the tool from offering to unprotect worksheets

Build 151: Added PowerQuery M code to Object search. Fixed a bug regarding finding cell references in Objects

Build 150: Added a Reset button on the settings form to reset the license registration

Build 149: You can now choose whether or not to display references in your formula more than once

Build 148: Fixed the wrong calculation setting (use 1 core only) of the add-in

Build 147: Improved the layout of the Object references dialog

Build 146: Fixed problem with not finding reference due to an alt+enter character in the formula

Build 145: Fixed problem with not finding reference to total row of tables in the format Table1[[#Totals],[May-2019]]

Build 144: VBA Project is now signed with a trusted code signing certificate

Build 143: Improved the Check Formulas interface

Build 142: Fixed a bug regarding handling of string literals in a formula

Build 141: Couple of bugs fixed

Build 140: Enabled selecting pivot table belonging to pivot chart

Build 139: Changed check for registration

Build 138: Fixed registration issue

Build 137: Fixed a bug with Visualization

Build 136: Internal build

Build 135: Fixed an issue with finding references in chart SERIES formula

Build 134: Improved finding, displaying and reporting of cell references used by Objects

Build 133: Internal build

Build 132: Updated the tool so it works with the new Excel Data Types (Geography and Stock) and with the new Dynamic Array references

Build 131: Improved display of Object dependencies, improved tiling windows next to Excel.

Build 129: Improved performance of formula checking and reporting significantly

Build 128: Added code to improve updating experience

Build 127: Adapted site addresses to use https

Build 126: Fixed runtime error due to third-party add-ins when checking if tool is installed

Build 125: Included Array formulas in circular reference checks

Build 124: Updated registration check

Build 122: Fixed a bug with finding Dependents introduced with build 121

Build 121: Improved performance of Dependents search

Build 120: Fixed small bug

Build 119: Added Report Function count

Build 118: Added option to settings to enable or disable automatic unhiding of worksheets

Build 117: Improved working of Tile option placing the dialog next to the Excel window, Improved reporting of formulas

Build 116: Enabled cyrrilic (and other non-western character sets)

Build 115: Fixed a small bug only occurring when you have a chartsheet selected

Build 114: Improved performance on analysing conditional formatting formulas

Build 113: Finally fixed an intermittent issue with the Visualize functionality on Excel 2016

Build 112: Fixed progress bar problems by enabling user to use taskbar instead

Build 111: Improved performance for Excel 2013 and 2016

Build 110: Fixed a crash of Excel due to corrupted add-in file

Build 109: Fixed repositioning of screen when re-activating dialog from Excel

Build 108: Added jump to pivotsource when you trace precedents when in a pivottable

Build 107: Fixed bug where range names were no longer listed (introduced with build 105)

Build 106: Fixed issue with short-cut keys not responding immediately after openening Excel

Build 105: Fixed bug, Pivottables were not listed when pointing to a table name

Build 104: Improved scrolling the selected cells into the viewable area of the screen

Build 103: Fixed not selecting off-sheet references from the precedents/dependents trees (introduced in build 102)

Build 102: Fixed bug changing objects hotkey; added work-around for people having issues with the Visualise option; fixed an issue with selecting objects from the treeview.

Build 101: Added workaround for Excel bug affecting the "Display Equation" function in non-English Excel versions

Build 100: Fixed bug regarding visualising merged cells

Build 099: Fixed small bug in report formulas

Build 098: Added "Display formulas as a mathematical equation"

Build 097: Improved scrolling within tables

Build 096: Added back double-click functionality to the treeview (same action as clicking Do ActiveCell)

Build 094: Fixed a bug causing crashes of Excel 2013 and 2016 when other certain add-ins are loaded

Build 093: Fixed a bug regarding unhiding worksheets

Build 092: Fixed a bug related to Excel 2013 and 2016; Improved Off-sheet references report to include checks for Tables and range names

Build 091: Variuos bug fixes

Build 088: Added a off-sheet references report which visualises the inter-sheet formula links of your workbook

Build 087: VBA code signed with new certificate, please make sure you update to this version!

Build 086: Skipped

Build 085: Fixed bug regarding setting errortracing hotkey

Build 084: Fixed objects bug and added automatic update of precedents/dependents tracking

Build 083: Fixed a bug in the sheet stats module

Build 082: Fixed a bug regarding whole-column references

Build 081: Improved error tracing, fixed bug regarding sheetnames with a pipe character

Build 080: Improved performance during finding dependents, improved Objects listing

Build 079: Fixed a number of windowing problems related to Excel 2013 and 2016, fixed a bug in the Find Object references

Build 078: Fixed hang when using check formulas

Build 077: Fixed asking for passwords when saving file

Build 076: Fixed runtime error when closing Excel with no workbooks present

Build 075: Added unprotecting/protecting of workbook and worksheets, Fixed issue with removal of arrows, fixed issue

Build 074: Not published

Build 073: Fixed tiny bug regarding editing of the formula in the dialog

Build 072: Improved window handling for Excel 2013 and up

Build 071: Fixed small issue with formula report and worksheet names that resemble dates

Build 069: Fixed bug (treeview not responding)

Build 068: Added Alt key support for hotkeys

Build 067: Avoid trying to remove arrows when no arrows have been added

Build 066: Improved grouping in the Objects dialog, fixed screenupdating bug

Build 065: Improved calculation timing

Build 064: Fixed a window resizing bug

Build 063: Fixed bug regarding startup screen not appearing in Excel 2013 (caused by an Office 2013 Update)

Build 062: Fixed bug regarding hotkeys not registering

Build 061: Fixed bug regarding startup screen not appearing in Excel 2013

Build 060: Added a Formula Report button

Build 059: fixed a compile error

Build 058: Changed registration behaviour, registered copies no longer access the Internet to check for registration

Build 057: Various small improvements

Build 056: Fixed bug regarding local range names.

Build 055: Added highlight formula blocks on "Check formulas" dialog, Improved performance of Check formulas.

Build 054: Enabled disable of hotkeys

Build 053: Improved handling of hotkeys

Build 052:

Added double-click to the treeviews

Build 051:

Added hotkey support for Objects and Check Formulas.

Build 050:

Fixed some bugs, improved reporting by adding hyperlinks

Build 049:

Improved progress bar for searching objects, prevented blank workbook when opening files from network shares

Build 048:

Made object search optional in find precedents/dependents

Build 047:

Beta version

Build 046:

Fixed a bug regarding Visualisation

Build 045:

Added the Object References feature to the find dependents and precedents dialog too.

Build 044:

Added the Object References feature

Build 043:

Fixed an issue with opening empty workbooks in 64 bit Excel 2013
Fixed issue with leaving an entry in the find dialog

Build 042:

Visualization: Improved placement of pictures of cells which are out of view

Build 041:

You can now change the visualisation colors in Settings.

Build 040:

Visualize option now also available for Excel 2003!

Build 039:

Bugfix

Build 038:

Bugfix

Build 037:

Added a new option to the tool: Visualize precedents. The precedents of a cell are visualised directly on the worksheet.

Build 036:

Fixed a bug on the reporting module

Build 035:

Improved tracing errors function

Build 034:

Fixed small windowing bug regarding Excel 2013

Build 033:

Important update: fixed some bugs in the multi-level precedents searching

Build 032:

Improved error tracing and functioning of the Stop button

Build 031:

Fixed windows covering each other in Excel 2013

Build 030:

RefTreeAnalyser no longer removes registration information when the add-in is unchecked.

Build 029:

Fixed a bug for Excel 2013, prevents messages being covered by the applications screens

Build 027:

Minor bugfix

Build 026:

Fixed bug regarding files opened in protected mode

Build 025:

Fixed bug with absolute cell references to multiple columns.

Build 024: 

Bugfix: improved unneeded scrolling when selecting cells.

Build 023: 

Minor bugfix

Build 022:

Improved returning to the selection prior to starting the tool.

Build 021:

Fixed an Excel 2007 issue.

Build 020:

A seperate add-in file was created for Excel versions prior to Excel 2007.

Build 019:

In some situations in Excel 2003 the "About" toolbar button showed the about screen of another add-in.

Build 018:

Fixed some problems regarding Excel 2003.

Build 017:

Fixed some problems regarding Excel 2003.

Build 015:

Fixed a tiny problem in About screen: RefTree does not properly show the registration state

Build 014:
Fixed a compile error that only Excel 2003 exhibits

Build 013:
Fixed a problem related to setting hotkeys
Made sure form position is remembered (and can be reset)

Build 012:
It is now possible to change the hotkeys.
I have also made sure that when an update is available, the help file is automatically updated too.
