# Best way to protect VBA code in Excel workbook or Add-In - VbaCompiler
Ref : https://vbacompiler.com/best-way-protect-vba-code

## Overview
In order to properly discuss the best way to protect VBA code, the “VBA code protection” term needs to be defined first, as well as the criteria of its efficiency.

All software authors want to avoid a source code leak, especially if they want to monetize their work. Thus, it makes sense to treat the term ‘VBA code protection’ as protection from accessing the VBA source code and protection from VBA code algorithm restoration.

To compare different VBA code protection methods, we need to have a measurement which directly shows the effectiveness of each method.

The best measurement would be economical—how many resources were used to produce VBA code and what is the cost of the restoration of this code or algorithms from the secured code. An effective and reliable way to protect VBA code should have the cost of recovering VBA source code from protected code significantly higher than the cost of creating the same VBA code from scratch.

It should become economically disadvantageous to recover VBA code in this case, because it is cheaper to create this VBA code from scratch than to recover it from the protected code.

If W is defined as the work hours it took to create the VBA code and X as the amount of work hours to crack this protection to get access to the protected VBA code or restore the VBA code algorithms, then the relation between these quantities gives us the quality of VBA protection:

X < W (or X/W < 1) –  means that cost of recovering source code is lower than developing the VBA code. This is low VBA code protection efficiency.

X = W (or X/W = 1) – means that cost of the recovering the source code is comparable to the cost of developing the VBA code. This is moderate protection efficiency.

X > W (or X/W > 1) – Cost of recovering the source code is higher than developing the VBA code. This is high protection efficiency.

![image](https://user-images.githubusercontent.com/60865708/194113691-87f9b2e7-2721-4fd1-a439-a91e7e32363d.png)

 You may consider VBA password as protection from accidental changes to the VBA code by the customer. Anybody can find ways on how to remove VBA Project protection on the Internet.

https://stackoverflow.com/questions/1026483/is-there-a-way-to-crack-the-password-on-an-excel-vba-project

Also, many cheap commercial tools are available on the market to remove the VBA password. Recovering the VBA code access in this case is automated and its cost may be considered as equal to zero (X = 0).

This method has low protection efficiency.

## Unviewable VBA Project
![image](https://user-images.githubusercontent.com/60865708/194113796-b97b6d8c-93d8-49ca-b239-7a5f77ba5d0a.png)

There is a way to make VBA Project unviewable by altering several bytes of the Excel workbook or Excel Add-In file in a HEX-editor (or programmatically). After such changes, the Excel VBA Project shows the “Project unviewable” message and blocks access to the VBA source code. But you need to understand that **such restrictions exist only in the Microsoft VBA editor**. There are several software products which allow you to see the VBA source code of the unviewable VBA project. One such software is open source [LibreOffice package](https://www.libreoffice.org/download/download/).

In most sophisticated cases of “Unviewable VBA” approaches the LibreOffice cannot reach the VBA code, but this  [can be resolved with simple manipulations](https://stackoverflow.com/a/67237347).

This method has a **low protection efficiency** (X < W)rating and may be considered as a way to protect VBA code from accidental changes of the VBA source code by the customer.

## VBA obfuscation

Source code obfuscation is defined as “the deliberate act of creating source or machine code that is difficult for humans to understand.” [https://en.wikipedia.org/wiki/Obfuscation_(software)](https://en.wikipedia.org/wiki/Obfuscation_(software))

Obfuscation of VBA source code includes changing names of methods, variables, and constants to random, difficult to read names, as well as removing comments and VBA code indenting to reduce understanding of the code.

![image](https://user-images.githubusercontent.com/60865708/194114007-44f34294-436a-4473-ad39-2bc22eca58f3.png)

In case of obfuscating, the structure of the algorithm is left unchanged and may be traced to recover the algorithms. There is existing software which allows to recover obfuscated VBA code formatting and increase the readability of the obfuscated source code. [https://rubberduckvba.com/](https://rubberduckvba.com/)

Simple features of any text editor such as “Find and Replace” lets you change obfuscated names to more readable and meaningful ones.

Practically, *VBA obfuscators* do not protect VBA code, because tracing of the code allows to recover all of the VBA source code logic.

So, in the case of **obfuscation the VBA code protection efficiency is low**. X < W (definition of X and W see above).

## Protect VBA code by Translating it to another programming language

The goal of this approach is to move VBA code logic into a DLL file and call DLL methods from VBA code.

This is the most efficient VBA code protection approach. Because the VBA source code is converted into the binary code of the EXE or DLL files.

The target language should be a compiled programming language, because any interpreting language (like VBA itself) doesn’t give effective protection.

The main drawbacks are the high cost and the error prone nature of this approach.

Consider the situation when the translation to a compiled programming language does not cover the whole VBA project code but only several VBA methods. In this case the X value, representing the amount of work hours, should be adjusted accordingly.

The cost of recovery of smaller, protected parts is obviously much less than the restoration of the VBA code as a whole. If the cost of re-writing such methods is less than recovering them from compiled modules then X gets this reduced cost and the efficiency of such protection goes down.

Below we consider the most popular languages for this approach—Visual Basic 6, .NET (C# or Visual Basic.NET), C/C++.

## VBA to Visual Basic 6 (VB6)

Visual Basic 6 (VB6) is an interpreted language but it also has the ability to compile VBA code into an EXE file or an ActiveX DLL.

The advantage of using this language lies in the simplicity of VBA to VB6 conversion. VB6 has the same syntax and semantics as VBA so you do not need to change a lot during code conversion.

Drawbacks:

VB6 doesn’t have a 64-bit version, so in case of creating an ActiveX Excel Add-In DLL it will only be possible to use the compiled DLL from Excel 32 bit.
 VB6 is an interpreting programming language, so all of its byte-code is saved inside the compiled EXE or DLL file. This means that even after compilation into an EXE or DLL file it may be decompiled into readable VB6 source code by VB-Decompiler.

[https://www.vb-decompiler.org/](https://www.vb-decompiler.org/)

So, with a ‘VB-Decompiler’ the protection code efficiency of this approach is reduced to the level of the VBA code obfuscation approach.

**Low protection efficiency** (X < W see above).

## VBA to VB.NET

In contrast to VB6, the .NET languages can create 32-bit as well as 64-bit versions of EXE and DLL files.

Converting VBA to .NET has a drawback in its architecture for resolving the VBA code protection tasks. It has a powerful ‘reflection’ mechanism which allows to convert the compiled code of .NET assembly into original source code. So, after the conversion of the VBA code to .NET it is possible to restore the source code from the created .NET assembly.

[https://www.red-gate.com/products/dotnet-development/reflector/](https://www.red-gate.com/products/dotnet-development/reflector/)

It is possible to apply code obfuscation to .NET assembly, but the efficiency of obfuscation has already been discussed above.

**Low protection efficiency** (X < W see above).

## VBA to C or C++

Translation of VBA code into C or C++ code provides very effective VBA protection. The source code restoration from compiled C/C++ EXE or DLL file of not a trivial project – is a very difficult task. In fact, it is so difficult and expensive that we can confidently say that it is practically impossible.

However, this approach has big drawback—C/C++ and VBA are very different programming languages. Conversion of VBA code to C/C++ is difficult and error prone, so the cost of such a conversion being not a trivial project is equal to the cost of creating the whole project from scratch.

**High protection efficiency** (X > W or X/W > 100 see above).

## Protect VBA code with VbaCompiler for Excel

[VbaCompiler for Excel](https://vbacompiler.com/) is VBA code protection software for Microsoft Excel. It converts the VBA source code to C language code and then compiles it into native Windows DLL. The efficiency of the VBA protection is the same as manual conversion of VBA to C/C++ language that was discussed above, but without the main drawbacks of the manual VBA to C/C++ conversion approach. **VBA compiler converts VBA code to DLL automatically, without the participation of a developer in the process**. You do not need to have any knowledge of C or C++ languages in order to use VBA compiler.

This means that the main drawback of the high cost of manual VBA to C/C++ conversion is eliminated.

**High VBA code protection efficiency** (X > W or X/W > 100 see above).

With VbaCompiler for Excel you have the best VBA code protection efficiency without the high cost of VBA to C/C++ code conversion work.

![image](https://user-images.githubusercontent.com/60865708/194114133-36a23797-520d-4379-81e3-42b333f45a03.png)

Beside the main VBA code protection feature – compilation into a DLL file the VbaCompiler for Excel provides more features to improve the VBA protection.

### Access control to Compiled VBA methods

Using  [the “Method expose mode” feature and the [DNXVBC_VBA_EXPOSED_METHOD] attribute](https://vbacompiler.com/vba-compiler-options/#methods_expose_mode) you may control what methods will be visible in the connective VBA code and exported from DLL file. So, you can remove some methods from the connective VBA code and from the DLL file export table. These methods will work inside the DLL module according to its internal calls, but will never be exposed outside of DLL.

### Encryption of all text literals

All VBA code text literals are removed from converted C-language source code during compilation. Text literals are encrypted by VbaCompiler and become available at run-time, only when the compiled workbook is started.
