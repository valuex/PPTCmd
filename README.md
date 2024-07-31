# PPTCmd - PPT Comand Palette
An add-in for MS PowerPoint to make PPT owns a Comand Palette like Visual Studio Code  
[中文说明](https://github.com/valuex/PPTCmd/blob/main/Readme_CN.md)

![image](https://github.com/user-attachments/assets/d9d16521-7dca-487e-83fe-7dac72765a28)  
Download: [Release](https://github.com/valuex/PPTCmd/releases)  

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
- **Id**: the **unique** index of the command, for internal indexing
- **GName**: the group name for the command,  for quick indexing by group name
- **EName**: the English name for the command
- **CName**: the Chinese name for the command
- **CmdType**: two types, `sys` - PPT internal command, `usr` - Macro command defined by user
- **Cmdlet**:
- - if `CmdType` is `sys`, the `Cmdlet` shall be exactly the same as listed in Column `Control Name` in `powerpointcontrols.xlsx` ;
- - if `CmdType` is `usr`, the `Cmdlet` shall be exactly the same as user defined macro's subroutine name in `MCMD.ppam` modules  ;
- **RTimes**: recoord the times that this command has been excuted, for sorting.
## **Cmdlet**-- PPT internal command
One can download the list of  PPT internal command from here  
https://www.microsoft.com/en-us/download/details.aspx?id=50745    
Basically, the `button` and  `toggleButton` type in  `Control Type` column of `powerpointcontrols.xlsx` can all be used in `CMDList.xml`.     
![image](https://github.com/user-attachments/assets/b39d4801-a22a-44c0-8219-a6b23b12a773)

## **Cmdlet**-- User Macro command
Right now only Chinese and Japanese tutorial of creating user macro in `*.ppam` file.  
Chinese: https://zhuanlan.zhihu.com/p/711155305  
Japanese: https://qiita.com/nkay/items/411ab09a0975aa48a449  
# Features
1. incremental search by PinYin  
2. use the arrow {Down} / {UP} to select the next / previous item  
3. use OpenBracket `[` / CloseBracket `]` to select the next / previous item.
