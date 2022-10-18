## Initial Setup
Also add EarlyBinding =1 or 0 into your VBA project Property conditional Compilation
![image](https://user-images.githubusercontent.com/60865708/196372852-4da94d37-f535-4d96-9b40-2f75290620ba.png)
 
### Step 1 : Create Class with If Else Block

```vba
Option Explicit

#If EarlyBinding = 1 Then
    Private m_dictionary As Dictionary
#Else
    Private m_dictionary As Object
#End If

Private Sub Class_Initialize()
    #If EarlyBinding = 1 Then
        Set m_dictionary = New Dictionary
    #Else
        Set m_dictionary = CreateObject("Scripting.Dictionary")
    #End If
End Sub

#If EarlyBinding = 1 Then
Public Property Get myDictionary() As Dictionary
 Set myDictionary = m_dictionary
End Property
#Else
Public Property Get myDictionary() As Object
 Set myDictionary = m_dictionary
#End If

End Property
```

### Step2 : Implement Class in to your Module

```vba
Public Sub Test()
    Dim dict As New clsDict
    dict.myDictionary.Add "Vivek", 100
    dict.myDictionary.Add "Deven", 80
    dict.myDictionary.Add "Druv", 60
    
    'Count
    Debug.Print "Total Items:" & dict.myDictionary.Count
    
    'Loop Using For
    Dim i As Long
    For i = 0 To dict.myDictionary.Count - 1
        Debug.Print dict.myDictionary.Keys()(i), dict.myDictionary.Items()(i)
    Next i
End Sub
```
