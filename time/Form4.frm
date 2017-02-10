VERSION 5.00
Begin VB.Form Form_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计时器"
   ClientHeight    =   2280
   ClientLeft      =   4905
   ClientTop       =   3885
   ClientWidth     =   7815
   DrawMode        =   8  'Xor Pen
   DrawStyle       =   2  'Dot
   FillStyle       =   0  'Solid
   Icon            =   "Form4.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MouseIcon       =   "Form4.frx":319A
   Palette         =   "Form4.frx":3A64
   PaletteMode     =   2  'Custom
   Picture         =   "Form4.frx":4225A6
   ScaleHeight     =   2280
   ScaleWidth      =   7815
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "直接关机"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   12
      Top             =   1750
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "只提醒"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5760
      TabIndex        =   11
      Top             =   1750
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "启动加密窗口"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2760
      TabIndex        =   10
      Top             =   1750
      Width           =   210
   End
   Begin VB.TextBox Ts 
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Text            =   "秒"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tm 
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Text            =   "分"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Th 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Text            =   "时"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "提醒"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   1750
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "启动屏锁精灵"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   1750
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "直接关机"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   1750
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "隐藏"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      MouseIcon       =   "Form4.frx":453F9A
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   3360
      MouseIcon       =   "Form4.frx":4540EC
      MousePointer    =   99  'Custom
      Picture         =   "Form4.frx":4549B6
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   7320
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "时"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   6120
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4245
   End
End
Attribute VB_Name = "Form_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
Check3.Value = 0
Check2.Enabled = False
Check3.Enabled = False
End If
If Check1.Value = 0 Then
Check2.Enabled = True
Check3.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Dim msg1, msg2, msg3
If Check1.Value = 1 Then
msg1 = Check1.Caption
Else
If Check2.Value = 1 Then
msg2 = Check2.Caption
End If
If Check3.Value = 1 Then
msg3 = Check3.Caption
End If
End If
msg = MsgBox("确定你设置的时间为：" & Text2.Text & "时" & Text3.Text & "分，" & "到时后的命令为：" & msg1 & msg2 & msg3, vbOKCancel)
If msg = 1 Then
Text2.Enabled = False
Text3.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command1.Enabled = False
End If
End Sub

Private Sub shape2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Form_GotFocus()
Me.Top = 0
End Sub

Private Sub Form_Load()




Timer1.Interval = 1000

Th.Text = Format(Now, "hh")
Tm.Text = Right(Format(Now, "hhmm"), 2)
Ts.Text = Format(Now, "ss")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Top < 0 Then
Me.Top = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Text2.Text = "" Or Text3.Text = "" Then
End
Else

l = Shell("c:\sgxt\ˇ36_pcms.exe", vbHide)

End
End If
End Sub



Private Sub Image1_Click()
If Text3.Text < 10 Then
Text3.Text = "0" & Text3.Text
End If
If Text2.Text < 10 Then
Text2.Text = "0" & Text2.Text
End If
Dim msg1, msg2, msg3
If Check1.Value = 1 Then
msg1 = Check1.Caption
Else
If Check2.Value = 1 Then
msg2 = Check2.Caption
End If
If Check3.Value = 1 Then
msg3 = Check3.Caption
End If
End If
msg = MsgBox("确定你设置的时间为：" & Text2.Text & "时" & Text3.Text & "分，" & "到时后的命令为：" & msg1 & msg2 & msg3, vbOKCancel)
If msg = 1 Then
Text2.Enabled = False
Text3.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command1.Enabled = False
End If
End Sub

Private Sub Label2_Click()
If Text2.Enabled = False Then
Me.Visible = False
End If

End Sub










Private Sub Text2_KeyPress(KeyAscii As Integer)

If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> Asc(".") And KeyAscii <> 8 Then
KeyAscii = 0
End If

End Sub

Private Sub Text3_Change()
If IsNumeric(Text3.Text) = False Then
Text3.Text = ""
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> Asc(".") And KeyAscii <> 8 Then
KeyAscii = 0
End If

End Sub


Private Sub Text2_Change()

If IsNumeric(Text2.Text) = False Then
Text2.Text = ""
End If

End Sub

Private Sub Timer1_Timer()
'窗体隐藏
If Me.Top <= 0 Then
Me.Top = -2570
End If

'时间累加
Ts.Text = Ts.Text + 1
If Ts.Text < 10 Then
Ts.Text = "0" & Ts.Text
End If
If Ts.Text = 60 Then
Ts.Text = "00"
Tm.Text = Tm.Text + 1
If Tm.Text < 10 Then
Tm.Text = "0" & Tm.Text
End If
End If
If Tm.Text = 60 Then
Tm.Text = "00"
Th.Text = Th.Text + 1
If Th.Text < 10 Then
Th.Text = "0" & Th.Text
End If
End If
If Th.Text = 24 Then
Th.Text = "00"
End If
Label1.Caption = "当前时间：" & Th.Text & "时" & Tm.Text & "分" & Ts.Text & "秒"


'判断是否到时
If Text2.Text = Th.Text Then
If Text3.Text = Tm.Text Then
If Check3.Value = 1 Then
s = Shell("c:\sgxt\ˇ36_pcms.exe", vbHide)
End If
If Check1.Value = 1 Then
Shell "cmd /c shutdown -s -f -t 1", vbHide
End If
If Check2.Value = 1 Then
Form1.Show
End If
End If
End If
End Sub




