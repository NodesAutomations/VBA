### Hyphen `-`
- Command Line version `-INSERT'
- Hyphen is required so we can use any autocad command without dialog box

### Semicolon `;`
- Semicolon is used as replacement for enter key press

### BackSlash `\`
- Pause for User input
- For example when we use circle command and it require centerpoint from user, so we can use blackslash
- this will pause current command so user can pick new point on active drawing

### Build Full Command to set Specific Layer 
- Full command : `^C^C -LAYER;S;TEXT;;`
- `^C^C` to cancel previous command
- `-LAYER` to start layer command without UI
- `;` to press enter, start layer command
- `S;` to set layer
- `TEXT;` name of layer then enter
- `;` Last enter to finish Command

### Run VBA Module
`
^C^C_-vbarun;Project.dvb!ExplodeBlocks
`
