### Overview
- Building custom Ribbon UI in AutoCAD is important for AutoCAD only Macro's
- With AutoCAD cui Client Can easliy use it's macro, without remembering any commands

### Steps to create new CUI
- We need two files for this
- DVB File: This contain vba project
- CUI File: This contain Ribbon UI Autocad

### DVB File
- You can create DVB file using AutoCAD VBA Editor
- Put DVB file into AutoCAD Support Folder in Programfiles for easy access 

### CUI File
- You can create CUI file Using AutoCAD CUI Editor
- You can keep CUI File whereever you like, You have to load it only once usign `CUILOAD` command and it will get automatically loaded whever you'll open new autoCAD
- to create new CUI file refer this [doc](https://github.com/NodesAutomations/VBA/blob/master/UI_AutoCAD/Custom%20Palettes.md)

### Project Refernce
- [Upwork_Marco_ExplodeBlocks](https://github.com/NodesAutomations/Upwork_Marco_ExplodeBlocks/releases/tag/v0.1.0)
