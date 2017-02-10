VERSION 5.00
Begin VB.Form juanzhu 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "赞助我们"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   Icon            =   "juanzhu.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   Picture         =   "juanzhu.frx":319A
   ScaleHeight     =   4005
   ScaleWidth      =   3840
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Text            =   """start page""=""http://www.2345.com/?3872"""
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "我愿意"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "您是否愿意赞助我们"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"juanzhu.frx":34B8E
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "  赞助我们不需要用户出一分钱，只需要将2345导航站设为主页就可以完成赞助我们的需求"
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"juanzhu.frx":34C22
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "首页"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      MouseIcon       =   "juanzhu.frx":34CC1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "赞助我们很简单："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "juanzhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
'Open App.Path & "\setpage.reg" For Output As #4
'Print #4, "REGEDIT4"
'Print #4, ""
'Print #4, "[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main]"
'Print #4, Text1.Text
'Close #4
Dim A As String
A = App.Path + "\" + "setpage.reg"
Shell "regedit.exe /s """ & A & """"
Shell "explorer.exe http://www.2345.com/?3872", 1
MsgBox ("感谢您的支持，您的支持使我们继续开发的最大动力")

End Sub



Private Sub Label4_Click()
Shell "explorer.exe http://120343.24la.com.cn/", 1
End Sub
