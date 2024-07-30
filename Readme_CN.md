# PPTCmd - PPT Comand Palette
PPT扩展， 实现类似于Visual Studio Code中的Comand Palette功能，即弹出面板，搜索并执行命令。  
下载： [Release](https://github.com/valuex/PPTCmd/releases)

# 使用方法
1. 在 `%appdata%` 目录下创建 `PPTCmd`文件夹  
2. 将下载的文件包中的 `CMDList.xml` 复制到 `PPTCmd`文件夹
3. 将下载的文件包中的 `MCMD.ppam` 复制到 `%appdata%\Microsoft\AddIns\` 
4. 双击 `setup.exe` 安装扩展 （需要在安装前关闭所有的PowerPoint）
5. 启动PowerPoint，点击按钮即可看到界面  
  ![image](https://github.com/user-attachments/assets/6f8990cf-6e6f-4c21-9543-e9250be5693a)


# 高级用法 Advanced Usage
目前在 `CMDList.xml`中只配置了一些demo命令, 用户可根据需要自行配置。

##  `CMDList.xml` 中的配置说明
`<cmd Id="1" GName="r" EName="RunCMD" CName="斜体" CmdType="sys" Cmdlet="Italic" RTimes="10"/>`
- **Id**: 命令编号, 留待后用(for making the listview more concise)
- **GName**: 命令的分组名, 留待后用 (可实现快速检索)
- **EName**: 命令的英文名
- **CName**: 命令的中文名
- **CmdType**: 分为两类, `sys` - PPT 内部命令, `usr` -用户创建的VBA 宏名  
- **Cmdlet**:
- - 如果 `CmdType` 是 `sys`,  `Cmdlet` **必须** 要跟 `powerpointcontrols.xlsx` 中的列 `Control Name` 中的单元格内容完全相同 ;
- - 如果 `CmdType` 是 `usr`,  `Cmdlet` **必须** 要跟用户在`MCMD.ppam`中定义的VBA宏名称完全相同 ;
- **RTimes**: 记录运行次数，用于后续排序.
### **Cmdlet** 中的PPT内部命令
可从如下地址下载 
https://www.microsoft.com/en-us/download/details.aspx?id=50745  
基本上，`powerpointcontrols.xlsx` 中的列【 `Control Type` 】为 `button` 和 `toggleButton` 的都可以配置到 `CMDList.xml`中.   
![image](https://github.com/user-attachments/assets/b39d4801-a22a-44c0-8219-a6b23b12a773)
### **Cmdlet** 中的用户自定义命令
可以到知乎上去看这篇文章  
[ppt 扩展（ppam ）开发](https://zhuanlan.zhihu.com/p/711155305)
# 特性
1. 支持拼音首字母检索命令。检索说明:
   - 只能匹配连续字符串，比如输入`abc`，可以检索到列表中的`abcd`，但检索不到列表中的`abdc`；
   - 输入不含空格的字符串，扩展会在命令中检索输入的字符串；
   - 输入`a空格b`，可以在分组名称中检索`a`，在命令中检索`b`。扩展只会把**第一个空格**前的内容用于检索分组名称。
   

# 开发背景
Office 2021后已经有了一个搜索框(Alt+Q)，但是不支持搜索宏命令，也不支持拼音首字母检索。
# 引用
1. NPinyinPro: https://www.nuget.org/packages/NPinyinPro
