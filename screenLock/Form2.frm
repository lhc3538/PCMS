VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "修改密码"
   ClientHeight    =   4440
   ClientLeft      =   4395
   ClientTop       =   4320
   ClientWidth     =   5205
   ForeColor       =   &H80000008&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      MouseIcon       =   "Form2.frx":0CCA
      TabCaption(0)   =   "修改密码"
      TabPicture(0)   =   "Form2.frx":15A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtReg"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "修改背景"
      TabPicture(1)   =   "Form2.frx":15C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "背景音乐"
      TabPicture(2)   =   "Form2.frx":15DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Image6"
      Tab(2).Control(2)=   "Image7"
      Tab(2).Control(3)=   "Text10"
      Tab(2).Control(4)=   "Check2"
      Tab(2).ControlCount=   5
      Begin VB.TextBox TxtReg 
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Text            =   """DisableTaskmgr""=dword:00000000"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "使用背景音乐"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   21
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73800
         TabIndex        =   20
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "背景设置"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   4575
         Begin VB.TextBox Text9 
            Height          =   270
            Left            =   480
            TabIndex        =   18
            Text            =   "Text9"
            Top             =   2040
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   4080
            TabIndex        =   17
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   1650
            ItemData        =   "Form2.frx":15F8
            Left            =   240
            List            =   "Form2.frx":15FA
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "多张背景"
            Height          =   255
            Left            =   1800
            TabIndex        =   15
            Top             =   860
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3960
            Top             =   2880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command2 
            Caption         =   "确定"
            Height          =   495
            Left            =   120
            MouseIcon       =   "Form2.frx":15FC
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            MousePointer    =   3  'I-Beam
            TabIndex        =   13
            Top             =   480
            Width           =   2655
         End
         Begin VB.Image Image5 
            Height          =   330
            Left            =   3000
            MouseIcon       =   "Form2.frx":1EC6
            MousePointer    =   99  'Custom
            Picture         =   "Form2.frx":2790
            Top             =   840
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Image Image4 
            Height          =   330
            Left            =   3750
            Picture         =   "Form2.frx":2D60
            Top             =   480
            Width           =   720
         End
         Begin VB.Image Image3 
            Height          =   375
            Left            =   1680
            MouseIcon       =   "Form2.frx":3372
            MousePointer    =   99  'Custom
            Picture         =   "Form2.frx":3C3C
            Top             =   3000
            Width           =   1200
         End
         Begin VB.Image Image1 
            Height          =   1695
            Left            =   720
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label4 
            Caption         =   "图片路径："
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   7
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Tag             =   "2"
         Text            =   "Text5"
         ToolTipText     =   "2"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确认修改"
         Height          =   495
         Left            =   360
         MouseIcon       =   "Form2.frx":4701
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   1
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   -73200
         MouseIcon       =   "Form2.frx":4FCB
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":511D
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Image Image6 
         Height          =   330
         Left            =   -70920
         MouseIcon       =   "Form2.frx":5A42
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":5B94
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "路径："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74520
         TabIndex        =   19
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   315
         Left            =   1920
         MouseIcon       =   "Form2.frx":61A6
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":6A70
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "重新确认密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "修改后密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "请输入原密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CheckApplicationIsRun(ByVal szExeFileName As String) As Boolean
On Error GoTo Err
Dim WMI
Dim Obj
Dim Objs
CheckApplicationIsRun = False
Set WMI = GetObject("WinMgmts:")
Set Objs = WMI.InstancesOf("Win32_Process")
For Each Obj In Objs
If InStr(UCase(szExeFileName), UCase(Obj.Description)) <> 0 Then
CheckApplicationIsRun = True
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
End If
Next
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
Err:
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
End Function
Private Sub Check1_Click()

If Check1.Value = 1 Then
Image1.Visible = False
List1.Visible = True
Image5.Visible = True
Else
Image1.Visible = True
List1.Visible = False
Image5.Visible = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
Text10.Enabled = False
Image6.Enabled = False
Else
Text10.Enabled = True
Image6.Enabled = True
End If

End Sub

Private Sub Command1_Click()
If Text1.Text = Text4.Text Then
If Text2.Text = Text6.Text Then
b = Len(Text2.Text)
For i = 1 To b
A = Asc(Mid(Text2.Text, i, 1))
b = b & "," & A * 66
Next i
Text5.Text = b

Open "c:\SGxt\MMtext\GJmm.dat" For Output As #2
Print #2, Text5.Text
Close #2
l = MsgBox("修改成功，下次使用生效.", vbOKOnly, "密码")
Unload Form1
Unload Form2
Unload Form3

Else
l = MsgBox("您输入的密码前后不一致，请重新输入！", vbOKOnly, "错误")
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
End If
Else
l = MsgBox("密码错误，请重新输入！", vbOKOnly, "错误")
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
End If
End Sub

Private Sub Command2_Click()
FileCopy Text7.Text, "c:\SGxt\ico\主界面." & Right(Text7.Text, 3)
msg = MsgBox("背景保存成功", vbOKOnly)
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
'重新开启任务栏
Shell "explorer.exe "
 If CheckApplicationIsRun("player.exe") = True Then '退出播放器
Shell "taskkill /im player.exe /f", vbHide
Else
Exit Sub
End If
'恢复任务管理器
Open App.Path & "\0.reg" For Output As #4
Print #4, "Windows Registry Editor Version 5.00"
Print #4, ""
Print #4, "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System]"
Print #4, TxtReg.Text
Close #4
Dim A As String
A = App.Path + "\" + "0.reg"
Shell "regedit.exe /s """ & A & """"
'删除临时文件
Kill (App.Path & "\0.reg")
Kill (App.Path & "\1.reg")

Open "c:\SGxt\MMtext\GJmm.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, A
Loop
Close #1
Text3.Text = A

l = Split(Text3.Text, ",")
For i = 1 To UBound(l)
Text4.Text = Text4.Text & Chr(l(i) / 66)
Next i

End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2


End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Image2_Click()
If Text1.Text = Text4.Text Then
If Text2.Text = Text6.Text Then
b = Len(Text2.Text)
For i = 1 To b
A = Asc(Mid(Text2.Text, i, 1))
b = b & "," & A * 66
Next i
Text5.Text = b

Open "c:\SGxt\MMtext\GJmm.dat" For Output As #2
Print #2, Text5.Text
Close #2
l = MsgBox("修改成功，下次使用生效.", vbOKOnly, "密码")
End


Else
l = MsgBox("您输入的密码前后不一致，请重新输入！", vbOKOnly, "错误")
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
End If
Else
l = MsgBox("密码错误，请重新输入！", vbOKOnly, "错误")
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
End If
End Sub

Private Sub Image3_Click()
If Check1.Value = 1 Then
For i = 1 To Text8.Text - 1
Text9.Text = List1.List(i)
FileCopy Text9.Text, "c:\SGxt\主界面" & i & ".jpg"
Next i
Open "c:\SGxt\ico\num.dat" For Output As #2
Print #2, Text8.Text
Close #2
msg = MsgBox("背景保存成功", vbOKOnly, "成功")
End

Else

If Text7.Text = "" Then
msg = MsgBox("请选择图片", vbOKOnly)
Else
FileCopy Text7.Text, "c:\SGxt\主界面." & Right(Text7.Text, 3)
msg = MsgBox("背景保存成功", vbOKOnly)
End
End If
End If
End Sub

Private Sub Image4_Click()
CommonDialog1.Filter = "Batch Files (*.jpg)|*.jpg" '
CommonDialog1.Action = 1
Text7.Text = CommonDialog1.FileName
End Sub

Private Sub Image5_Click()
If Right(Text7.Text, 3) = "jpg" Then
List1.List(Text8.Text) = Text7.Text
Text8.Text = Text8.Text + 1
Else
msg = MsgBox("只支持*.jpg格式的图片文件", vbOKOnly, "格式")
End If
End Sub

Private Sub Image6_Click()
CommonDialog1.Filter = "Batch Files (*.mp3)|*.mp3" '
CommonDialog1.Action = 1
Text10.Text = CommonDialog1.FileName
End Sub

Private Sub Image7_Click()
If Check2.Value = 1 Then
FileCopy Text10.Text, "c:\sgxt\music\BM.mp3"
msg = MsgBox("背静音乐保存成功。", vbOKOnly, "操作成功")
Else
Kill ("c:\sgxt\music\BM.mp3")
End If


End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Clipboard.Clear
End Sub

Private Sub Text7_Change()
If Image1.Visible = True Then

Image1.Picture = LoadPicture(Text7.Text)
End If
End Sub
