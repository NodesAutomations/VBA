### Add Custom Ribbon
- Download [RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) to Add Custom Ribbon xml to any xlsm/docm/pptm File
- Add Custom Tab using Below Template
- Add Icon PNG files 64px

## Xml Template 
### Custom Tab
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"> 
   <ribbon> 
     <tabs> 
        <tab id="CustomTab" label="TrainingAutomation" keytip="M"> 
         <group id="SampleGroup" label="Formatting"> 
           <button id="SplitButton" label="&amp;Split" image="Splitter" size="large" onAction="SplitShape" /> 
           <button id="MergeButton" label="&amp;Merge" image="Merge" size="large" onAction="MergeShape" /> 
         </group > 
       </tab> 
     </tabs> 
   </ribbon> 
 </customUI>
```
### Open Custom Ribbon On workbook Load
```xml
<customUI onLoad="RibbonOnLoad" xmlns="http://schemas.microsoft.com/office/2009/07/customui"> 
   <ribbon> 
     <tabs> 
        <tab id="CustomTab" label="Shoeb Lakhi" keytip="S"> 
         <group id="SampleGroup" label="DataBase"> 
           <button id="SyncButton" label="&amp;Sync" imageMso="SynchronizeHtml" size="large" onAction="SyncData" />           
         </group > 
       </tab> 
     </tabs> 
   </ribbon> 
 </customUI>
```
Add This To VBA Module
```vba
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Dim ribRibbon As IRibbonUI
    Set ribRibbon = ribbon
    ribRibbon.ActivateTab ("CustomTab")
End Sub
```
Special Symbols
```
Ampersand - & [& #38;]
Forward Slash - / [& #47;]
Backward Slash - \ [& #92;]
New Line - [& #13;]
Apostrophe - ' [& apos;]
```


## Attribute Referance
- Taken From [Ribbon Attributes](https://bettersolutions.com/vba/ribbon/tab.htm)
- [Github Repo:Office Identifiers](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
- [Microsoft Doc](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee691833(v=office.14)?redirectedfrom=MSDN)

## Image Gallary
- [Bert Image Gallary](https://bert-toolkit.com/imagemso-list.html)

### Tags
| Tag  | Detail |
| ------------- | ------------- |
|enabled	|Specifies whether the control is enabled or not.|
|id|	Specifies the identifier for a custom control. All custom controls must have unique identifiers.|
|idMso|	Specifies the identifier of a built-in control.|
|idQ|	Specifies a qualified identifier for a control.|
|insertAfterMso|	Specifies the identifier of a built-in control that this control should be inserted after.|
|insertAfterQ	|Specifies the qualified identifier of a control that this control should be inserted after.|
|insertBeforeMso|	Specifies the identifier of a built-in control that this control should be inserted before.|
|insertBeforeQ	|Specifies the qualified identifier of a control that this control should be inserted before.|
|keytip|	Specifies a string to be used as the keytip for this control.|
|label|	Specifies a string to be used as the label for this control.|
|screentip	|Specifies a string to be used as the supertip for this control.|
|showImage|	Specifies whether this control should display its image.|
|showLabel|	Specifies whether this control should display its label.|
|supertip|	Specifies a string to be used as the supertip for this control.|
|tag	|Specifies an arbitrary string that can be used to hold data or identify the control.|
|visible	|Specifies whether the control is visible or not.|


