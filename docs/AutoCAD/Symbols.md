### Hyphen `-`
- Command Line version `-INSERT'
- Hyphen is required so we can use any autocad command without dialog box

### Semicolon `;`
- Semicolon is used as replacement for enter key press

### BackSlash `\`
- Pause for User input
- For example when we use circle command and it require centerpoint from user, so we can use blackslash
- this will pause current command so user can pick new point on active drawing

### AT `@`
- Use Last used point

### Asterik `*`
- Repeats the macro until cancelled by user

### Build Full Command to set Specific Layer 
- Full command : `^C^C -LAYER;S;TEXT;;`
- `^C^C` to cancel previous command
- `-LAYER` to start layer command without UI
- `;` to press enter, start layer command
- `S;` to set layer
- `TEXT;` name of layer then enter
- `;` Last enter to finish Command
- You can keep going as long as you want, you don't need to use hyphen again for next command

### Run VBA Module
`
^C^C_-vbarun;Project.dvb!ExplodeBlocks
`
### Reference
- [About Special Control Characters in Command Macros](https://knowledge.autodesk.com/support/autocad-lt/learn-explore/caas/CloudHelp/cloudhelp/2018/ENU/AutoCAD-LT/files/GUID-DDDB6E26-75E1-4643-8C6A-BEAEBA83A424-htm.html)
