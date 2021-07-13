### Read Xml File Using MSXML
- Add Reference to Microsoft XML, v3.0
```vba
Public Sub LoadDocument()
    Dim xDoc As MSXML2.DOMDocument
    Set xDoc = New MSXML2.DOMDocument
    xDoc.validateOnParse = False
    If xDoc.Load("C:\Users\Ryzen2600x\Downloads\Point.xml") Then
        ' The document loaded successfully.
        ' Now do something intersting.
        DisplayNode xDoc.ChildNodes, 0
    Else
        ' The document failed to load.
        ' See the previous listing for error information.
    End If
End Sub

Public Sub DisplayNode(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
   ByVal Indent As Integer)

   Dim xNode As MSXML2.IXMLDOMNode
   Indent = Indent + 2

   For Each xNode In Nodes
      If xNode.NodeType = NODE_TEXT Then
         Debug.Print Space$(Indent) & xNode.ParentNode.nodeName & _
            ":" & xNode.NodeValue
      End If

      If xNode.HasChildNodes Then
         DisplayNode xNode.ChildNodes, Indent
      End If
   Next xNode
End Sub
```





### Resources
- [Loop through Nodes](https://excel-macro.tutorialhorizon.com/vba-excel-read-xml-by-looping-through-nodes/)
- [Update XML File](https://excel-macro.tutorialhorizon.com/vba-excel-update-xml-file/)
- [Read Data From XML](https://excel-macro.tutorialhorizon.com/vba-excel-read-data-from-xml-file/)
