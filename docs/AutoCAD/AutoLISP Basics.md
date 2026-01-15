# AutoLISP Basics

## Resources
- [AutoLISP Beginner Tutorials](https://www.afralisp.net/autolisp/tutorials/index.php?category_id=1)
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

## Create Custom Command
- Below is an example of a custom command that prompts the user to select two points and then
- `c:test` is the name of the command you will type in AutoCAD to run this function.
```lisp
(defun c:test ()
    ;define the function
 
	(setq a (getpoint "\nEnter First Point : "))
	;get the first point
 
	(setq b (getpoint "\nEnter Second Point : "))
	;get the second point
 
	(command "Line" a b "")
	;draw the line
 
        (princ)
        ;clean running
 
)	;end defun
(princ)
;clean loading
```

## Function for Calculations
- You can also define functions that are not directly callable as commands.
```lisp
(defun myFunction (x y)
  (+ x y)
)
(princ (myFunction 5 10)) ; This will print 15
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

### GetDist
- Used to get a distance from the user.
```lisp
(setq userDistance (getdist "\nEnter a distance: "))
(princ userDistance) ; This will print the entered distance
```

### Polar
- Used to calculate a new point based on a starting point, distance, and angle.
```lisp
(setq startPoint (getpoint "\nSelect a starting point: "))
(setq newPoint (polar startPoint (/ pi 4) 10)) ; 45 degrees, 10 units
(princ newPoint) ; This will print the new point coordinates
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
