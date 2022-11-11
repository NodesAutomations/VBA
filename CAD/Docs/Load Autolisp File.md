### Steps to Load Auto lisp File
- This Proces is same for AutoCAD and ZWCAD only change in UI
- Just use Appload Command to load autolisp 
- use related command name to trigger command, in case of below script use `cut`

### Sample Lisp file
```lisp
(defun c:cut(/ p1 p2 p3 p4 pt1 pt2)
(if (not def) (setq def 10.0))
(setvar "osmode" (+ 0))
(initget 1)
(setq pt1 (getpoint "\nPick first point to draw cut line:"))
(initget 1)
(setq pt2 (getpoint pt1 "\nPick second point to draw cut line:"))
;|(setvar "osmode" 0)|;
(setq dis (getdist (strcat "\nPick or Enter Cut gap <" (rtos def 2 4) ">:")))

(if (eq dis nil) (setq dis def) (setq def dis))

(setq
     p1  (polar pt1 (angle pt2 pt1) dis)
     p2  (polar p1 (angle pt1 pt2)
         (+ (/ (distance pt1 pt2) 2.0) (/ (* 0.5 dis) 2.0) )  )
     p3  (polar p2 (+ (angle pt1 pt2) (dtr 75)) (* 1.5 dis))
     p5  (polar p2 (angle pt1 pt2) (* 1.5 dis))
     p4  (polar p5 (- (angle pt1 pt2) (dtr 105)) (* 1.5 dis))
     p6  (polar pt2 (angle pt1 pt2) dis)
)
(setvar "cmdecho" 0)
(command ".PLINE" p1 "W" "0" "0" p2 p3 p4 p5 p6 "")
(setvar "cmdecho" 1)
(princ)
)
     
;*************************************************************************
(defun dtr(a) 
(* pi (/ a 180.0))
)
;*************************************************************************
(defun rtd(a) 
(* a (/ 180.0 pi))
)
;*************************************************************************
(princ)
```
