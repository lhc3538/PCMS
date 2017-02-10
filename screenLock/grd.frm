VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "其它"
   ClientHeight    =   7215
   ClientLeft      =   -60
   ClientTop       =   -15
   ClientWidth     =   6000
   Icon            =   "grd.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   1680
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1080
      TabIndex        =   8
      Text            =   "120343.24la.com.cn"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   3735
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   6588
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
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   5400
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Text            =   "pp,stop"
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton Command4 
         Caption         =   "计时器"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   960
         MouseIcon       =   "grd.frx":1043E
         MousePointer    =   99  'Custom
         Picture         =   "grd.frx":10D08
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdlook 
         Caption         =   "开机日志"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   960
         MouseIcon       =   "grd.frx":119D2
         MousePointer    =   99  'Custom
         Picture         =   "grd.frx":1229C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "版权所有"
      Height          =   975
      Left            =   5280
      MouseIcon       =   "grd.frx":12F66
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   180
      Left            =   5280
      TabIndex        =   5
      Top             =   1920
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   0
      Picture         =   "grd.frx":13830
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   5160
      MouseIcon       =   "grd.frx":14053
      MousePointer    =   99  'Custom
      Picture         =   "grd.frx":1491D
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdlook_Click()
l = Shell("c:\sgxt\look.exe", vbHide)
End

End
End Sub

Private Sub Command1_Click()

frmAbout.Show 1
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Command4_Click()
l = Shell("c:\sgxt\jsq.exe", vbHide)

End
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Web1.Navigate Text3.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False

End Sub


Private Sub Image1_Click()
Me.Visible = False
End Sub

Private Sub Timer1_Timer()
Text4.Text = Text4.Text + 1
If Text4.Text = "300" Then
Form5.Show 1
End If

End Sub
