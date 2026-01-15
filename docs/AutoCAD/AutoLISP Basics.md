# AutoLISP Basics

## Resources
- [Beginner AutoLISP Tutorial](https://www.youtube.com/playlist?list=PLwqMD67__ep39iFjt1vBLs7MJMbKfnHz0)

## Basics
- Use `VLIDE` to write and test AutoLISP code.
- You Can Open Existing Lisp file by `File` -> `Open File` 
- Windows
  - Main Code Window : This is where you write and edit your AutoLISP code.
  - Console/Command Line : This is where you can see output and error messages.
  - Build Output Window : This shows the results of building or compiling your code.
  - Trace Window : This is useful for debugging, showing the flow of execution.
- You can Use `Appload` button or `Ctrl + F9` to load your lisp code into AutoCAD.
- You can add your Lisp file to Startup Suite so that it loads automatically when AutoCAD starts.
- You can also run Lisp Code Directly in AutoCAD Command Line 

## Sample Code 
- Here is a simple AutoLISP function that prints "Hello, World!" to the command line when you type `hello` in AutoCAD.
- `defun` defines a new function, `c:` indicates it's a command, and `princ` is used to print text.
```lisp
(defun c:hello ()
  (princ "Hello, World!")
)
```

## Basic Syntax

### Comments
- Use `;` to add comments in your code. Anything after `;` on the same line is ignored by AutoLISP.
```lisp
; This is a comment
(defun c:example ()
  (princ "This is an example function") ; This prints a message
)
```

### Print
- Used to display messages in the command line.
```lisp
(Princ "Hello, AutoLISP World!")
```

### Message Box
- Used to display a message box to the user.
```lisp
(alert "This is a message box!")
```

### If Statement
- Used for conditional execution of code.
```lisp
(if (> 10 5)
  (princ "10 is greater than 5")
  (princ "10 is not greater than 5")
)
```

### IF Else Statement
- Used for conditional execution with an alternative path.
```lisp
(if (> 10 5)
  (princ "10 is greater than 5")
  (princ "10 is not greater than 5")
)
```

### Loops
- Used to repeat a block of code multiple times.
```lisp
(repeat 5
  (princ "This will print 5 times")
)
```

## Special Functions

### Setq
- Used to assign values to variables.
```lisp
(setq myVariable 10)
(princ myVariable) ; This will print 10
```

### GetPoint
- Used to get a point from the user in the drawing area.
```lisp
(setq userPoint (getpoint "\nSelect a point: "))
(princ userPoint) ; This will print the selected point coordinates
```

### Command 
- Used to execute AutoCAD commands from within AutoLISP.
- `""` is used to indicate the end of a command sequence, similar to pressing Enter.
```lisp
(setq a (getpoint "\nEnter First Point : "))
(setq b (getpoint "\nEnter Second Point : "))
(Command "LINE" a b "")
```

### Table Search
- Used to search for a value in a list or table.
```lisp
(tblsearch "value" '("value1" "value2" "value3"))
```
