### Start PowerPoint With Specific Template
```vba
 Dim templatePath As String
    templatePath = "C:\Users\Ryzen2600x\Downloads\Test.potx"
      
    
    'PowerPoint Object Variables
    Dim PPApp  As PowerPoint.Application
    Dim PPPresentation  As PowerPoint.Presentation
    
    'Open PowerPoint Application
    Set PPApp = CreateObject("PowerPoint.Application")
    PPApp.WindowState = ppWindowMinimized
    
    'Create PPT With Specific Template
    Set PPPresentation = PPApp.Presentations.Open(templatePath, False, True, True)
```

### Start PowerPoint with Specific PPT File
```vba
  Dim templatePath As String
    templatePath = "C:\Users\Ryzen2600x\Downloads\Sample.pptx"
      
    
    'PowerPoint Object Variables
    Dim PPApp  As PowerPoint.Application
    Dim PPPresentation  As PowerPoint.Presentation
    
    'Open PowerPoint Application
    Set PPApp = CreateObject("PowerPoint.Application")
    PPApp.WindowState = ppWindowMinimized
    
    'Create PPT With Specific Template
    Set PPPresentation = PPApp.Presentations.Open(templatePath, False, True, True)
    
    
    PPPresentation.SaveAs ("C:\Users\Ryzen2600x\Downloads\Test.pptx")
```
