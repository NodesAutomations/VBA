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
- ref : https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey

### for numpad keys
```vba
Sub ReAssignKeypad()
Application.OnKey "{096}", "KeyPad0"
Application.OnKey "{097}", "KeyPad1"
Application.OnKey "{098}", "KeyPad2"
Application.OnKey "{099}", "KeyPad3"
Application.OnKey "{100}", "KeyPad5"
Application.OnKey "{101}", "KeyPad5"
Application.OnKey "{102}", "KeyPad6"
Application.OnKey "{103}", "KeyPad7"
Application.OnKey "{104}", "KeyPad8"
Application.OnKey "{105}", "KeyPad9"
Application.OnKey "{106}", "KeyPadMult"
Application.OnKey "{107}", "KeyPadPlus"
Application.OnKey "{109}", "KeyPadMinus"
Application.OnKey "{110}", "KeyPadPoint"
Application.OnKey "{111}", "KeyPadDiv"
End Sub
```
