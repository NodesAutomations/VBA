### so there's multiple way to assign shortcut to macr
- manually use can assign assign shortcut key by using macros>Options
- usign ribbon xml 
- using vba code, you can run this code on Workbook_open() event in thisworkbook
```
Application.OnKey "^b", "MyProgram"
```
or
```
Application.MacroOptions macro:="MyProgram", Description:="Description of the Macro", _
hasshortcutkey:=True, ShortcutKey:="^b"
```
- ref : https://vmlogger.com/excel/2013/06/assign-a-shortcut-key-using-excel-vba/
