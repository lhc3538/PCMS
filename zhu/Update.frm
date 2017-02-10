VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Update 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "软件更新"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5115
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5115
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "下载地址"
      Height          =   1695
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   4335
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "天空软件站下载【推荐】"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MouseIcon       =   "Update.frx":319A
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "华军软件站下载【推荐】"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "非凡软件站下载"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "多特软件站下载"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   5
      Text            =   "http://120343.24la.com.cn/shengji.html"
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      ExtentX         =   2990
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "优软官方网站"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客服QQ：353850101        软件QQ群：111082285"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "最新版本："
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "v1.8.160"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "当前版本："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "私人计算机管理系统   程序更新"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
      
       Web1.Navigate Text1.Text
End Sub

Private Sub Label10_Click()
Shell "explorer.exe http://www.skycn.com/soft/59274.html", 1
End Sub

Private Sub Label5_Click()
Shell "explorer.exe http://120343.24la.com.cn/", 1
End Sub

Private Sub Label6_Click()
Shell "explorer.exe http://www.crsky.com/soft/19447.html", 1
End Sub

Private Sub Label7_Click()
Shell "explorer.exe http://www.duote.com/soft/23208.html", 1
End Sub

Private Sub Label9_Click()
Shell "explorer.exe http://www.newhua.com/soft/106147.htm", 1
End Sub
