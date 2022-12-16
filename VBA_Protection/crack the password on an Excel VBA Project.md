# Is there a way to crack the password on an Excel VBA Project? - Stack Overflow

Reference : https://stackoverflow.com/questions/1026483/is-there-a-way-to-crack-the-password-on-an-excel-vba-project?rq=1

You can try this direct `VBA` approach which doesn't require HEX editing. It will work for any files (*.xls, *.xlsm, *.xlam ...).

Tested and works on:

> Excel 2007 Excel 2010 Excel 2013 - 32 bit version Excel 2016 - 32 bit version
> 

Looking for 64 bit version? See [this answer](https://stackoverflow.com/a/31005696/4342479)

## How it works

I will try my best to explain how it works - please excuse my English.

1. The VBE will call a system function to create the password dialog box.
2. If user enters the right password and click OK, this function returns 1. If user enters the wrong password or click Cancel, this function returns 0.
3. After the dialog box is closed, the VBE checks the returned value of the system function
4. if this value is 1, the VBE will "think" that the password is right, hence the locked VBA project will be opened.
5. The code below swaps the memory of the original function used to display the password dialog with a user defined function that will always return 1 when being called.

## Using the code

**Please backup your files first!**

1. Open the file(s) that contain your locked VBA Projects
2. Create a new xlsm file and store this code in **Module1** 
    
    `code credited to Siwtom (nick name), a Vietnamese developer`
    
    ```
    Option Explicit
    
    Private Const PAGE_EXECUTE_READWRITE = &H40
    
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (Destination As Long, Source As Long, ByVal Length As Long)
    
    Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
            ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
    
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
            ByVal lpProcName As String) As Long
    
    Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
            ByVal pTemplateName As Long, ByVal hWndParent As Long, _
            ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    
    Dim HookBytes(0 To 5) As Byte
    Dim OriginBytes(0 To 5) As Byte
    Dim pFunc As Long
    Dim Flag As Boolean
    
    Private Function GetPtr(ByVal Value As Long) As Long
        GetPtr = Value
    End Function
    
    Public Sub RecoverBytes()
        If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
    End Sub
    
    Public Function Hook() As Boolean
        Dim TmpBytes(0 To 5) As Byte
        Dim p As Long
        Dim OriginProtect As Long
    
        Hook = False
    
        pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    
        If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
    
            MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
            If TmpBytes(0) <> &H68 Then
    
                MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6
    
                p = GetPtr(AddressOf MyDialogBoxParam)
    
                HookBytes(0) = &H68
                MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
                HookBytes(5) = &HC3
    
                MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
                Flag = True
                Hook = True
            End If
        End If
    End Function
    
    Private Function MyDialogBoxParam(ByVal hInstance As Long, _
            ByVal pTemplateName As Long, ByVal hWndParent As Long, _
            ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
        If pTemplateName = 4070 Then
            MyDialogBoxParam = 1
        Else
            RecoverBytes
            MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                               hWndParent, lpDialogFunc, dwInitParam)
            Hook
        End If
    End Function
    
    ```
    
3. Paste this code under the above code in **Module1** and run it
    
    ```
    Sub unprotected()
        If Hook Then
            MsgBox "VBA Project is unprotected!", vbInformation, "*****"
        End If
    End Sub
    
    ```
    
4. Come back to your VBA Projects and enjoy.
