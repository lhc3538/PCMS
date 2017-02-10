VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "私人计算机管理系统 卸载向导"
   ClientHeight    =   5205
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   7695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7695
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7695
      Begin VB.Image Image1 
         Height          =   4695
         Left            =   -120
         Picture         =   "Form1.frx":C84A
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   2490
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "输入用户密码后，单击[下一步]继续。"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "在开始卸载之前，请确认 私人计算机管理系统 并未运行当中。"
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "这个程序将全程指引您 私人计算机管理系统 的卸载进程。 "
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "欢迎使用 私人计算机管理系统  卸载向导"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "请输入用户密码："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2760
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   7680
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   6360
      MouseIcon       =   "Form1.frx":1137B
      Picture         =   "Form1.frx":114CD
      Top             =   4750
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   4560
      MouseIcon       =   "Form1.frx":11A0D
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":11B5F
      Top             =   4750
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

Open "c:\SGxt\MMtext\GJmm.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, temp
Loop
Close #1
Text3.Text = temp
l = Split(Text3.Text, ",")
For i = 1 To UBound(l)
Text4.Text = Text4.Text & Chr(l(i) / 66)
Next i

End Sub

Private Sub Image2_Click()
If Text1.Text = Text4.Text Then
Text1.Text = ""
Form1.Visible = False
Form2.Visible = True
Else
msg = MsgBox("密码错误，您无权卸载！", vbOKOnly, "密码错误")
End If
End Sub

Private Sub Image3_Click()
End
End Sub

