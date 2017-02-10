VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "修改密码"
   ClientHeight    =   4515
   ClientLeft      =   4395
   ClientTop       =   4320
   ClientWidth     =   5205
   ForeColor       =   &H80000008&
   Icon            =   "z.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
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
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      MouseIcon       =   "z.frx":0CCA
      TabCaption(0)   =   "修改密码"
      TabPicture(0)   =   "z.frx":15A4
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
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "个性化"
      TabPicture(1)   =   "z.frx":15C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "背景设置"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   4575
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3720
            Top             =   960
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command2 
            Caption         =   "确定"
            Height          =   495
            Left            =   120
            MouseIcon       =   "z.frx":15DC
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   1080
            MousePointer    =   3  'I-Beam
            TabIndex        =   13
            Top             =   480
            Width           =   2655
         End
         Begin VB.Image Image4 
            Height          =   330
            Left            =   3750
            Picture         =   "z.frx":1EA6
            Top             =   480
            Width           =   720
         End
         Begin VB.Image Image3 
            Height          =   315
            Left            =   1680
            MouseIcon       =   "z.frx":24B8
            MousePointer    =   99  'Custom
            Picture         =   "z.frx":2D82
            Top             =   3000
            Width           =   1125
         End
         Begin VB.Image Image1 
            Height          =   1695
            Left            =   720
            Stretch         =   -1  'True
            Top             =   1080
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
         Height          =   735
         Left            =   600
         MouseIcon       =   "z.frx":417C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
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
      Begin VB.Image Image2 
         Height          =   315
         Left            =   1920
         MouseIcon       =   "z.frx":4A46
         MousePointer    =   99  'Custom
         Picture         =   "z.frx":5310
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
Private Sub Command1_Click()
If Text1.Text = Text4.Text Then
If Text2.Text = Text6.Text Then
b = Len(Text2.Text)
For i = 1 To b
a = Asc(Mid(Text2.Text, i, 1))
b = b & "," & a * 66
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
Dim a As String
Open "c:\SGxt\MMtext\GJmm.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, a
Loop
Close #1
Text3.Text = a

l = Split(Text3.Text, ",")
For i = 1 To UBound(l)
Text4.Text = Text4.Text & Chr(l(i) / 66)
Next i

End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2
Unload Form3

End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Image2_Click()
If Text1.Text = Text4.Text Then
If Text2.Text = Text6.Text Then
b = Len(Text2.Text)
For i = 1 To b
a = Asc(Mid(Text2.Text, i, 1))
b = b & "," & a * 66
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

Private Sub Image3_Click()
If Text7.Text = "" Then
msg = MsgBox("请选择图片", vbOKOnly)
Else
FileCopy Text7.Text, "c:\SGxt\ico\主界面." & Right(Text7.Text, 3)
msg = MsgBox("背景保存成功", vbOKOnly)
End If
End Sub

Private Sub Image4_Click()
CommonDialog1.Filter = "Batch Files (*.jpg)|*.jpg" '
CommonDialog1.Action = 1
Text7.Text = CommonDialog1.FileName
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Clipboard.Clear
End Sub

Private Sub Text7_Change()
Image1.Picture = LoadPicture(Text7.Text)
End Sub
