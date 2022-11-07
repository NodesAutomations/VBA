### Basis Syntax

```vba
Public Enum Options
    Option1
    Option2
    Option3
End Enum

Public Sub Example()
    Debug.Print Options.Option1 'Prints 0
    Debug.Print Options.Option2 'Prints 1
    Debug.Print Options.Option3 'Prints 2
End Sub
```

- Enums are Accesible at Project Level
- If you want to access any enum in another module you can use `vbaProjectName.EnumName.Choice`
