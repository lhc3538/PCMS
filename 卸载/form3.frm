VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "私人计算机管理系统 卸载向导"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   Icon            =   "form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7695
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   2400
      TabIndex        =   0
      Top             =   -240
      Width           =   5415
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "谢谢合作"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "   请把您的使用本软件的意见反馈给我们，是对我们的最大支持，您的意见是我们最宝贵的东西！"
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "问题反馈"
         Height          =   255
         Left            =   4200
         MouseIcon       =   "form3.frx":319A
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "完成卸载，感谢您的支持"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Label Label5 
      Caption         =   "http://youchuang.uueasy.com/read.php?tid=12"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   4920
      MouseIcon       =   "form3.frx":32EC
      MousePointer    =   99  'Custom
      Picture         =   "form3.frx":343E
      Top             =   4750
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   7680
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   -120
      Picture         =   "form3.frx":3B18
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Dim Browser As Object

                  url = "http://youchuang.uueasy.com/read.php?tid=12"

                 Set Browser = CreateObject("InternetExplorer.Application")

                  Browser.Visible = True

                  Browser.Navigate (url)

Open App.Path & "\kill.bat " For Append As #1
Print #1, "del   " & App.EXEName & ".exe "
Print #1, "del   kill.bat "
Print #1, "del   %0"
Close #1
Shell App.Path & "\Kill.bat ", 0
End
End Sub

Private Sub Label2_Click()
Dim Browser As Object

                  url = "http://youchuang.uueasy.com/read.php?tid=12"

                 Set Browser = CreateObject("InternetExplorer.Application")

                  Browser.Visible = True

                  Browser.Navigate (url)

Open App.Path & "\kill.bat " For Append As #1
Print #1, "del   " & App.EXEName & ".exe "
Print #1, "del   kill.bat "
Print #1, "del   %0"
Close #1
Shell App.Path & "\Kill.bat ", 0
End
End Sub
