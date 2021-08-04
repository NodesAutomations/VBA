## Read Xml File Using MSXML
- Add Reference to Microsoft XML, v6.0

### Sample XML File for Testing
<details>
<summary>Click to toggle contents of `code`</summary>
  Add Here Whatever YOu want
</details>

```xml
<?xml version="1.0"?>
<catalog>
   <book id="bk101">
      <author>Gambardella, Matthew</author>
      <title>XML Developer's Guide</title>
      <genre>Computer</genre>
      <price>44.95</price>
      <publish_date>2000-10-01</publish_date>
      <description>An in-depth look at creating applications 
      with XML.</description>
   </book>
   <book id="bk102">
      <author>Ralls, Kim</author>
      <title>Midnight Rain</title>
      <genre>Fantasy</genre>
      <price>5.95</price>
      <publish_date>2000-12-16</publish_date>
      <description>A former architect battles corporate zombies, 
      an evil sorceress, and her own childhood to become queen 
      of the world.</description>
   </book>
   <book id="bk103">
      <author>Corets, Eva</author>
      <title>Maeve Ascendant</title>
      <genre>Fantasy</genre>
      <price>5.95</price>
      <publish_date>2000-11-17</publish_date>
      <description>After the collapse of a nanotechnology 
      society in England, the young survivors lay the 
      foundation for a new society.</description>
   </book>
   <book id="bk104">
      <author>Corets, Eva</author>
      <title>Oberon's Legacy</title>
      <genre>Fantasy</genre>
      <price>5.95</price>
      <publish_date>2001-03-10</publish_date>
      <description>In post-apocalypse England, the mysterious 
      agent known only as Oberon helps to create a new life 
      for the inhabitants of London. Sequel to Maeve 
      Ascendant.</description>
   </book>
   <book id="bk105">
      <author>Corets, Eva</author>
      <title>The Sundered Grail</title>
      <genre>Fantasy</genre>
      <price>5.95</price>
      <publish_date>2001-09-10</publish_date>
      <description>The two daughters of Maeve, half-sisters, 
      battle one another for control of England. Sequel to 
      Oberon's Legacy.</description>
   </book>
   <book id="bk106">
      <author>Randall, Cynthia</author>
      <title>Lover Birds</title>
      <genre>Romance</genre>
      <price>4.95</price>
      <publish_date>2000-09-02</publish_date>
      <description>When Carla meets Paul at an ornithology 
      conference, tempers fly as feathers get ruffled.</description>
   </book>
   <book id="bk107">
      <author>Thurman, Paula</author>
      <title>Splish Splash</title>
      <genre>Romance</genre>
      <price>4.95</price>
      <publish_date>2000-11-02</publish_date>
      <description>A deep sea diver finds true love twenty 
      thousand leagues beneath the sea.</description>
   </book>
   <book id="bk108">
      <author>Knorr, Stefan</author>
      <title>Creepy Crawlies</title>
      <genre>Horror</genre>
      <price>4.95</price>
      <publish_date>2000-12-06</publish_date>
      <description>An anthology of horror stories about roaches,
      centipedes, scorpions  and other insects.</description>
   </book>
   <book id="bk109">
      <author>Kress, Peter</author>
      <title>Paradox Lost</title>
      <genre>Science Fiction</genre>
      <price>6.95</price>
      <publish_date>2000-11-02</publish_date>
      <description>After an inadvertant trip through a Heisenberg
      Uncertainty Device, James Salway discovers the problems 
      of being quantum.</description>
   </book>
   <book id="bk110">
      <author>O'Brien, Tim</author>
      <title>Microsoft .NET: The Programming Bible</title>
      <genre>Computer</genre>
      <price>36.95</price>
      <publish_date>2000-12-09</publish_date>
      <description>Microsoft's .NET initiative is explored in 
      detail in this deep programmer's reference.</description>
   </book>
   <book id="bk111">
      <author>O'Brien, Tim</author>
      <title>MSXML3: A Comprehensive Guide</title>
      <genre>Computer</genre>
      <price>36.95</price>
      <publish_date>2000-12-01</publish_date>
      <description>The Microsoft MSXML3 parser is covered in 
      detail, with attention to XML DOM interfaces, XSLT processing, 
      SAX and more.</description>
   </book>
   <book id="bk112">
      <author>Galos, Mike</author>
      <title>Visual Studio 7: A Comprehensive Guide</title>
      <genre>Computer</genre>
      <price>49.95</price>
      <publish_date>2001-04-16</publish_date>
      <description>Microsoft Visual Studio 7 is explored in depth,
      looking at how Visual Basic, Visual C++, C#, and ASP+ are 
      integrated into a comprehensive development 
      environment.</description>
   </book>
</catalog>
```
### Sample Code TO Read XML and Remove First Child Node
```vba
Sub XMLTest()
    Dim xmlPath As String
    xmlPath = "C:\Users\Ryzen2600x\Downloads\Test - Copy.xml"
    
    Dim xDoc As New MSXML2.DOMDocument60
    
    'Load xml
    If Not xDoc.Load(xmlPath) Then
        Exit Sub
    End If
    
    'Loop Through All Nodes
    Call DisplayNode(xDoc.ChildNodes, 2)
    
    'Find Specific Node
    Dim xNode As MSXML2.IXMLDOMNode
    Set xNode = xDoc.SelectSingleNode("catalog").ChildNodes(0)
    xNode.ParentNode.RemoveChild xNode
    xDoc.Save (xmlPath)
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
### Code Snippet to Select Specific Node
```vba
    Dim presentationPrNode As MSXML2.IXMLDOMNode
    Set presentationPrNode = xDoc.SelectSingleNode("p:presentationPr")

    Dim clrMruNode As MSXML2.IXMLDOMNode
    Set clrMruNode = xDoc.SelectSingleNode("p:presentationPr/p:clrMru")
```

### Resources
- [Parse XML using VBA](https://stackoverflow.com/questions/11305/how-to-parse-xml-using-vba)
- [Microsoft Doc:Beginner's Guide XML DOM](https://docs.microsoft.com/en-us/previous-versions/aa468547(v=msdn.10)?redirectedfrom=MSDN)
- [Loop through Nodes](https://excel-macro.tutorialhorizon.com/vba-excel-read-xml-by-looping-through-nodes/)
- [Update XML File](https://excel-macro.tutorialhorizon.com/vba-excel-update-xml-file/)
- [Read Data From XML](https://excel-macro.tutorialhorizon.com/vba-excel-read-data-from-xml-file/)
