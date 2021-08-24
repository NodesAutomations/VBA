## Conditional Compilation
For conditional compilation first we need to define local or global conditional variable
### project level Variable
![image](https://user-images.githubusercontent.com/60865708/130563646-89b05e10-7468-4341-9529-f7a114765d09.png)

### Module Level
```vba
' Declare public compilation constant in Declarations section. 
#Const conDebug = 1 
```

### Use Case
```vba
Sub SelectiveExecution() 
 #If conDebug = 1 Then 
 . ' Run code with debugging statements. 
 . 
 . 
 #Else 
 . ' Run normal code. 
 . 
 . 
 #End If 
End Sub
```
