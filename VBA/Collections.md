### Use Case
- Collections are normally perfer when you don't know size 
- One Additional advantages is that you can store any object in collection as you like, including custom one

### Create new Collecton
```vba
    ' Declare
    Dim coll As New Collection

    ' Add item - VBA looks after resizing
    coll.Add "Apple"
    coll.Add "Pear"

    ' remove item - VBA looks after resizing
    coll.Remove 1
```
### Addin or Accesing Items using Key
```vba
' https://excelmacromastery.com/
Sub UseKey()

    Dim collMark As New Collection

    collMark.Add 45, "Bill"
    collMark.Add 67, "Hank"
    collMark.Add 12, "Laura"
    collMark.Add 89, "Betty"

    ' Print Betty's marks
    Debug.Print collMark("Betty")

    ' Print Bill's marks
    Debug.Print collMark("Bill")

End Sub
```
### Loop through Collection
```vba
Sub UserCollection()

    ' Declare and Create collection
    Dim collFruit As New Collection

    ' Add items
    collFruit.Add "Apple"
    collFruit.Add "Pear"
    collFruit.Add "Plum"

    ' Print all items
    Dim i As Long
    For i = 1 To collFruit.Count
        Debug.Print collFruit(i)
    Next i

End Sub
```
