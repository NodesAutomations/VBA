### How To See Add-in Code
Powerpoint add-in don't behave like Excel add-in you have to make some changes in windows Registary to Make Add-in Code Visible Follow Below Steps
- Open Registary Editor
- Navigate to the following key in the registry : HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options
- Find the key name DebugAddins and Set it's value to 1
  - If Key is not found create new DWORD32bit key named DebugAddins
- That's it save registrary file and now you'll be able to see powerpoint vba Project in VBA Code Explorer
