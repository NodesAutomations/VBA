### Ribbon with inbuilt Functions
- Remove OnLoad Event if you're making addin and don't want to activate your custom ribbon onload
- You need name of control to add button or group
- You can find Ribbon Control Id from Customize Ribbon Menu, Id name are written in round brackets
- In Below Screenshot, `BordersGallery` is inbuilt control

![RibbonControlID](/UI_Office/RibbonControlD.png)

```xml
<customUI onLoad="RibbonOnLoad"
    xmlns="http://schemas.microsoft.com/office/2009/07/customui">
    <ribbon>
        <tabs>
            <tab id="NodesVBAHelperTab" label="NodesVBAHelper" keytip="G">
               <group idMso="GroupArrange">
                    <button idMso="AlignObjects" />
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
```
### Tab and Group
```xml
<tab id="CustomTab" label="TrainingAutomation" keytip="M"> 
  <group id="SampleGroup" label="Formatting"> 
     <!--Add your button code here  -->
  </group> 
</tab> 
```
### Special Symbols for labels
```
Ampersand - & [& #38;]
Forward Slash - / [& #47;]
Backward Slash - \ [& #92;]
New Line - [& #13;]
Apostrophe - ' [& apos;]
```
### Simple Button
```xml
   <button id="AboutButton" label="About" imageMso="InformationDialog" size="large" onAction="AboutButton_Click" getScreentip="AboutButton_Tip" />
```
### Simple Button with custom Icon
- use png icons
- use 24 size for larger icons
- use 16 size for smaller icons
- Use white `#FFFFFF` for Main Color
- Use Red `ff3f3f` for Secondary color
  
### Split Button





### Tags
| Tag             | Detail                                                                                           |
| --------------- | ------------------------------------------------------------------------------------------------ |
| enabled         | Specifies whether the control is enabled or not.                                                 |
| id              | Specifies the identifier for a custom control. All custom controls must have unique identifiers. |
| idMso           | Specifies the identifier of a built-in control.                                                  |
| idQ             | Specifies a qualified identifier for a control.                                                  |
| insertAfterMso  | Specifies the identifier of a built-in control that this control should be inserted after.       |
| insertAfterQ    | Specifies the qualified identifier of a control that this control should be inserted after.      |
| insertBeforeMso | Specifies the identifier of a built-in control that this control should be inserted before.      |
| insertBeforeQ   | Specifies the qualified identifier of a control that this control should be inserted before.     |
| keytip          | Specifies a string to be used as the keytip for this control.                                    |
| label           | Specifies a string to be used as the label for this control.                                     |
| screentip       | Specifies a string to be used as the supertip for this control.                                  |
| showImage       | Specifies whether this control should display its image.                                         |
| showLabel       | Specifies whether this control should display its label.                                         |
| supertip        | Specifies a string to be used as the supertip for this control.                                  |
| tag             | Specifies an arbitrary string that can be used to hold data or identify the control.             |
| visible         | Specifies whether the control is visible or not.                                                 |


## Attribute Referance
- Taken From [Ribbon Attributes](https://bettersolutions.com/vba/ribbon/tab.htm)
- [Github Repo:Office Identifiers](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
- [Microsoft Doc](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee691833(v=office.14)?redirectedfrom=MSDN)