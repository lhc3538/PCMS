VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "实时资讯"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Text            =   "http://120343.24la.com.cn/adver3.html"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser Web3 
      Height          =   2295
      Left            =   6480
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   4048
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
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Text            =   "http://120343.24la.com.cn/adver2.html"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser Web2 
      Height          =   2895
      Left            =   6480
      TabIndex        =   3
      Top             =   0
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   5106
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   6360
      TabIndex        =   2
      Top             =   0
      Width           =   15
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "http://120343.24la.com.cn/资讯.html"
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      ExtentX         =   11033
      ExtentY         =   9763
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
   Begin VB.Label Label1 
      Caption         =   "    MM盛夏时刻"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6480
      MouseIcon       =   "Form5.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Web1.Navigate Text1.Text
Web2.Navigate Text2.Text
Web3.Navigate Text3.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload Form1
Unload Form2
Unload Form3
Unload Form4

End Sub

Private Sub Label1_Click()
Shell "explorer.exe http://120343.24la.com.cn/", 1
End Sub


