### Add Custom Ribbon


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
