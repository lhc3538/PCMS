VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "成功安装"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8220
   Icon            =   "finish.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8220
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "添加桌面快捷方式"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "开机启动"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   2640
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.Image Image3 
      Height          =   3735
      Left            =   -120
      Picture         =   "finish.frx":319A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2130
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   6960
      MouseIcon       =   "finish.frx":7CCB
      MousePointer    =   99  'Custom
      Picture         =   "finish.frx":7E1D
      Top             =   4850
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   5265
      Left            =   0
      Picture         =   "finish.frx":8586
      Top             =   0
      Width           =   8220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub RecurseTree(CurrPath As String)

    Dim sFileName As String

    Dim newPath As String

    Dim sPath As String

    Static oldPath As String

    sPath = CurrPath & "\"

    sFileName = Dir(sPath, 31) '31的含义∶31=vbNormal+vbReadOnly+vbHidden+vbSystem+vbVolume+vbDirectory

    Do While sFileName <> ""

    If sFileName <> "." And sFileName <> ".." Then

    If GetAttr(sPath & sFileName) And vbDirectory Then '如果是目录和文件夹

    newPath = sPath & sFileName

    RecurseTree newPath

    sFileName = Dir(sPath, 31)

    Else

    SetAttr sPath & sFileName, vbNormal

    Kill (sPath & sFileName)


    sFileName = Dir

    End If

    Else

    sFileName = Dir

    End If

    DoEvents

    Loop

    SetAttr CurrPath, vbNormal

    RmDir CurrPath

 

    End Sub

Private Sub Form_Load()
If Dir("c:\sgxt\", vbDirectory) = "" Then
 
 Else
 Check1.Value = 1
 Check1.Enabled = True
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Image2_Click
Shell App.Path & "\cell.exe", vbNormalFocus
End Sub

Private Sub Image2_Click()
If Dir("c:\sgxt\", vbDirectory) = "" Then
'未注册过执行的命令
MkDir "c:\SGxt"
MkDir "c:\SGxt\MMtext"
MkDir "c:\SGxt\looktext"
MkDir "c:\SGxt\ico"
MkDir "c:\SGxt\music"

FileCopy App.Path & "\kjjm.exe", "c:\sgxt\kjjm.exe"
FileCopy App.Path & "\lock.exe", "c:\sgxt\lock.exe"
Name "c:\sgxt\kjjm.exe" As "c:\sgxt\ˇ36_pcms.exe"
FileCopy App.Path & "\player.exe", "c:\sgxt\player.exe"



FileCopy App.Path & "\mima.dat", "c:\SGxt\MMtext\GJmm.dat"
FileCopy App.Path & "\pp.dat", "c:\SGxt\MMtext\pp.dat"
FileCopy App.Path & "\num.dat", "c:\SGxt\ico\num.dat"



FileCopy App.Path & "\ico\t.jpg", "c:\SGxt\ico\t.jpg"
FileCopy App.Path & "\ico\f.jpg", "c:\SGxt\ico\f.jpg"
FileCopy App.Path & "\ico\lock.ico", "c:\SGxt\ico\lock.ico"




FileCopy App.Path & "\SkinH_VB6.dll", "c:\SGxt\SkinH_VB6.dll"
FileCopy App.Path & "\skinH.she", "c:\SGxt\skinH.she"

FileCopy App.Path & "\ico\主界面.jpg", "c:\SGxt\主界面0.jpg"

Kill App.Path & "\mima.dat"


FileCopy App.Path & "\time\jsq.exe", "c:\sgxt\jsq.exe"
FileCopy App.Path & "\sj.dat", "c:\SGxt\looktext\sj.dat"
FileCopy App.Path & "\gs.dat", "c:\SGxt\looktext\gs.dat"
FileCopy App.Path & "\look\look.exe", "c:\sgxt\look.exe"
FileCopy App.Path & "\cell.exe", "c:\sgxt\cell.exe"
'判断是否写入开机启动项
If Check1.Value = 1 Then
Dim n
n = Shell(App.Path & "\注册.bat")
End If


If Check3.Value = 1 Then
Set WshShell = CreateObject("Wscript.shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oMyShortcut = WshShell.CreateShortcut(strDesktop + "\屏锁精灵.lnk") '此处为快捷名称
oMyShortcut.IconLocation = "c:\SGxt\ico\lock.ico" '此处为快捷图标
oMyShortcut.TargetPath = "C:\SGxt\lock.exe" '此处为源文件
oMyShortcut.Hotkey = "ALT+CTRL+C" ''此处为快捷热键
oMyShortcut.Save
End If








Unload Me


'----------------------------------------------------------------------------------------------

Else



'注册过执行的命令


FileCopy "c:\SGxt\MMtext\GJmm.dat", App.Path & "\mima.dat"
Dim strPathName As String
    strPathName = "c:\sgxt\"
  RecurseTree strPathName
MkDir "c:\SGxt"
MkDir "c:\SGxt\MMtext"
MkDir "c:\SGxt\looktext"
MkDir "c:\SGxt\ico"
MkDir "c:\SGxt\music"

FileCopy App.Path & "\kjjm.exe", "c:\sgxt\kjjm.exe"
FileCopy App.Path & "\lock.exe", "c:\sgxt\lock.exe"
Name "c:\sgxt\kjjm.exe" As "c:\sgxt\ˇ36_pcms.exe"
FileCopy App.Path & "\player.exe", "c:\sgxt\player.exe"



FileCopy App.Path & "\mima.dat", "c:\SGxt\MMtext\GJmm.dat"
FileCopy App.Path & "\pp.dat", "c:\SGxt\MMtext\pp.dat"
FileCopy App.Path & "\num.dat", "c:\SGxt\ico\num.dat"


FileCopy App.Path & "\ico\t.jpg", "c:\SGxt\ico\t.jpg"
FileCopy App.Path & "\ico\f.jpg", "c:\SGxt\ico\f.jpg"
FileCopy App.Path & "\ico\lock.ico", "c:\SGxt\ico\lock.ico"
FileCopy App.Path & "\SkinH_VB6.dll", "c:\SGxt\SkinH_VB6.dll"
FileCopy App.Path & "\skinH.she", "c:\SGxt\skinH.she"


FileCopy App.Path & "\ico\主界面.jpg", "c:\SGxt\主界面0.jpg"

Kill App.Path & "\mima.dat"


FileCopy App.Path & "\time\jsq.exe", "c:\sgxt\jsq.exe"
FileCopy App.Path & "\sj.dat", "c:\SGxt\looktext\sj.dat"
FileCopy App.Path & "\gs.dat", "c:\SGxt\looktext\gs.dat"
FileCopy App.Path & "\look\look.exe", "c:\sgxt\look.exe"
FileCopy App.Path & "\cell.exe", "c:\sgxt\cell.exe"
'判断是否写入开机启动项
If Check1.Value = 1 Then
n = Shell(App.Path & "\注册.bat")
End If



If Check3.Value = 1 Then
Set WshShell = CreateObject("Wscript.shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oMyShortcut = WshShell.CreateShortcut(strDesktop + "\屏锁精灵.lnk") '此处为快捷名称
oMyShortcut.IconLocation = "C:\SGxt\ico\lock.ico" '此处为快捷图标
oMyShortcut.TargetPath = "C:\SGxt\lock.exe" '此处为源文件
oMyShortcut.Hotkey = "ALT+CTRL+C" ''此处为快捷热键
oMyShortcut.Save
End If





Unload Me


End If
'--------------------------------------------------------------------------------------
End Sub
