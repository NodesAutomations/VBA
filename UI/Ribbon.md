### Add Custom Ribbon
- Download [RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) to Add Custom Ribbon xml to any xlsm/docm/pptm File
- Add Custom Tab using Below Template
- Add Icon PNG files 64px

### Xml Template
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
## Attribute Referance
- Taken From [Ribbon Attributes](https://bettersolutions.com/vba/ribbon/tab.htm)
### Tags
| Content Cell  | Content Cell  |
| Content Cell  | Content Cell  |


