# AutoLISP Basics

## Resources
- [Beginner AutoLISP Tutorial](https://www.youtube.com/playlist?list=PLwqMD67__ep39iFjt1vBLs7MJMbKfnHz0)

## Basics
- Use `VLIDE` to write and test AutoLISP code.

## Sample Code 
- Here is a simple AutoLISP function that prints "Hello, World!" to the command line when you type `hello` in AutoCAD.
- `defun` defines a new function, `c:` indicates it's a command, and `princ` is used to print text.
```lisp
(defun c:hello ()
  (princ "Hello, World!")
)
```

 