# For setting up AutoCAD Start Up macro 

### Using Project DVB Files
- You can store your autocad vba projects in DVB files
- Now If you want to use this macro on any drawing you have to manually load this dvb file everytime you open autocad
- But you can set it up to run it automatically
- First type `APPLOAD` command , it will open new windowd, on that windows there's group called startup suite, click on contents button there which will let you add dvb files for automatically load at startup
- Refer : [Automatically load dvb file](http://www.lee-mac.com/autoloading.html)

![image](https://user-images.githubusercontent.com/60865708/195878362-73314a6b-c59d-4e30-b124-c14705228e98.png)

### Using AutoLisp File

Upon opening a drawing or starting a new drawing, AutoCAD will search all listed support paths including the working directory for a file with the filename: ACADDOC.lsp. If one or more such files are found, AutoCAD will proceed to load the first file found.

With this knowledge one can edit or create an ACADDOC.lsp to include any AutoLISP expressions to be evaluated upon startup.

Things get a little complicated should there exist more than one ACADDOC.lsp file, so, to check if such a file already exists, type or copy the following line to the AutoCAD command line:
`
(findfile "ACADDOC.lsp")
`
Should a filepath be returned, in the steps that follow, navigate to this file and amend its contents. Else, if the above line returns nil, you can create your own ACADDOC.lsp in Notepad or the VLIDE and save it in an AutoCAD Support Path.
One clear advantage to using the ACADDOC.lsp to automatically load programs is that, upon migration, it may easily be copied from computer to computer, or indeed reside on a network to load programs on many computers simultaneously.

### Code to Assign Command for Project.dvb in AutoCAD Support Folder
```lisp
(defun c:Test() (command "vbarun" "Blocks.ExplodeBlocks"))
```

![image](https://user-images.githubusercontent.com/60865708/195880462-858884f9-894e-4bc1-8972-531f1bead006.png)

![image](https://user-images.githubusercontent.com/60865708/195880839-a8c76d20-735e-4988-9848-41c37b587cf0.png)

