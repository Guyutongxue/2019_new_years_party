# 2019_new_years_party

### 2018年班级联欢晚会PPT后台系统

## 总述

这是为2019年班级新年联欢晚会而制作的在线实时更新系统。其原理大体类似[2018年班级联欢晚会PPT后台系统](https://github.com/Guyutongxue/2018_new_year_party)，但又有些许更新和改进。

## 更新与改进

### 关于 cmd.exe

自从2018年4月起，学校禁用直接运行 cmd.exe 的权限，此政策使得无法再通过 VBA 调用系统 cmd 脚本（但仍可以双击运行），导致原有核心代码失效。好在学校保留了运行 PowerShell.exe 的权限，惟需将相关代码转写为 PowerShell 语言即可。

#### 更改 PowerShell 执行政策

PowerShell默认禁止运行任何脚本，此时需要在运行系统前手动执行 `Run_This_First.cmd` 脚本来更改用户执行政策，启用脚本运行。

##### Run_This_First.cmd

```BAT
@echo off
echo Setting PowerShell Execution Policy ...
powershell.exe -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned"
echo Done.
pause
```

#### PowerShell 更新脚本

如同去年的 `update.cmd` ，今年采用 `update.ps1`。代码如下：

##### update.ps1

```PowerShell
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptLocation
./wget64.exe -O info.txt https://guyutongxue.github.io/2019_new_years_party/info.txt
```

### 关于 VBA

#### 针对 PowerShell 的一些更改

将去年的 `update()` 函数作如下更改：

```VB
Sub UpdateData()
    Shell ("powershell.exe -File " + Application.ActivePresentation.Path + "\update.ps1")
End Sub
```

### 针对睡眠的一些更改

将计时器实现的睡眠改为调用 Windows API 实现的睡眠。

```VB
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Sub Sleep(millisecond As Integer)
    Dim Savetime As Double
    Savetime = timeGetTime
    While timeGetTime < Savetime + millisecond
        DoEvents
    Wend
End Sub
```

#### 针对 UTF-8 编码文件读取的一些更改

去年采用网上调用系统API读取的代码，今年改用 ADODB.Stream 组件直接读取，减少了代码量。

```VB
Function ReadFromTextFile(filepath As String)
    Dim str
    Set stm = CreateObject("adodb.stream")
    stm.CharSet = "UTF-8"
    stm.Open
    stm.LoadFromFile filepath
    str = stm.ReadText
    stm.Close
    Set stm = Nothing
    ReadFromTextFile = str
End Function
```

#### 针对等待 wget64.exe 运行的一些更改

将数据下载到本地需要一定的时间，去年的系统中，采用了全部等待 2s 再执行后续操作的策略。今年改用查询 WMI 服务循环监测 PowerShell.exe 进程来决定何时继续执行。下附全部核心代码。

```VB
Function IsWgetRunning()
    Dim strComputer  As String
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim i As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'powershell.exe'")
    If colProcessList.Count > 0 Then
    'MsgBox "aa"
        IsWgetRunning = True
    Else
        IsWgetRunning = False
    End If
End Function

Sub SwitchSlide()
    Call UpdateData
    Dim isLoop As Boolean
    isLoop = IsWgetRunning()
    While isLoop = True
        Call Sleep(100)
        isLoop = IsWgetRunning()
    Wend
    Dim text As String
    Dim strArr As Variant
    'text = UTF8_Decode(Application.ActivePresentation.Path + "\info.txt")
    text = ReadFromTextFile(Application.ActivePresentation.Path + "\info.txt")
    strArr = Split(text, Chr(10))
    If strArr(0) = "text" Then
        Application.ActivePresentation.Slides(2).Shapes("Title").TextFrame.TextRange.text = strArr(1)
        Application.ActivePresentation.Slides(2).Shapes("Text").TextFrame.TextRange.text = ""
        Dim i As Integer
        For i = 2 To UBound(strArr) - 1
            Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text = Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text + strArr(i) + Chr(10)
        Next
        Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text = Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text + strArr(UBound(strArr))
        Application.SlideShowWindows(1).View.GotoSlide (2)
    ElseIf strArr(0) = "slide" Then
        Application.SlideShowWindows(1).View.GotoSlide (Val(strArr(1)))
    End If
End Sub

```
### 关于后台终端

首先采用完美越狱的 iPad 作为后台终端，更加稳定。
其次增设 `ctrl.sh` ，制作一键操控的界面。

##### ctrl.sh

```Shell
#!/bin/bash
while true
do
	echo -e "PPTCtrl> \c"
	read cmd
	case $cmd in
	"c")
	./create.sh
	;;
	"s")
	./switch.sh
	;;
	"u")
	./upload.sh
	;;
	"q")
	exit 0
	;;
	*)
	echo "Command error."
	;;
	esac
done

```
其它脚本 `create.sh` 、 `upload.sh` 、 `switch.sh` 无大更新，不做赘述。