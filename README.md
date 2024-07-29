# PPTCmd - PPT Comand Palette
An add-in for MS PowerPoint to make PPT owns a Comand Palette like Visual Studio Code  
[中文说明](https://github.com/valuex/PPTCmd/blob/main/Readme_CN.md)
# Usage
1. Create foler `PPTCmd` under `%appdata%`
2. Copy `CMDList.xml` to the folder `PPTCmd`
3. Copy `MCMD.ppam` to `%appdata%\Microsoft\AddIns\` 
4. Double-click `setup.exe` to install the add-in
# Advanced Usage
Now only some demo commands are configed in  `CMDList.xml`, the user can open it and do the setting by himself / herself.  
Users can visit this page to download the internal commands list in PowerPoint.   
https://www.microsoft.com/en-us/download/details.aspx?id=50745  
Basically, all these command with the `Control Type` as `button` and `toggleButton` can be used in `CMDList.xml`.   
![image](https://github.com/user-attachments/assets/b39d4801-a22a-44c0-8219-a6b23b12a773)
## Explaination of `CMDList.xml`
`<cmd Id="1" GName="r" EName="RunCMD" CName="斜体" CmdType="sys" Cmdlet="Italic" RTimes="10"/>`
- **Id**: the index of the command, for future usage (for making the listview more concise)
- **GName**: the group name for the command, for future usage (quick indexing by group name)
- **EName**: the English name for the command
- **CName**: the Chinese name for the command
- **CmdType**: two types, `sys` - PPT internal command, `usr` - Macro command defined by user
- **Cmdlet**:
- - if `CmdType` is `sys`, the `Cmdlet` shall be exactly the same as listed in Column `Control Name` in `powerpointcontrols.xlsx` ;
- - if `CmdType` is `usr`, the `Cmdlet` shall be exactly the same as user defined macro's subroutine name in `MCMD.ppam` modules  ;
- **RTimes**: recoord the times that this command has been excuted, for sorting in the future.

# Features
1. incremental search by PinYin
