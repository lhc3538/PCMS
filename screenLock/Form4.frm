VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "其它"
   ClientHeight    =   2340
   ClientLeft      =   -60
   ClientTop       =   -15
   ClientWidth     =   5985
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
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
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      Begin VB.CommandButton Command4 
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
         MouseIcon       =   "Form4.frx":1043E
         MousePointer    =   99  'Custom
         Picture         =   "Form4.frx":10D08
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdlook 
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
         Left            =   0
         MouseIcon       =   "Form4.frx":1191D
         MousePointer    =   99  'Custom
         Picture         =   "Form4.frx":121E7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "计时器"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "开机日志"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "版权所有"
      Height          =   975
      Left            =   4200
      MouseIcon       =   "Form4.frx":12961
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   180
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   0
      Picture         =   "Form4.frx":1322B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   5160
      MouseIcon       =   "Form4.frx":13A4E
      MousePointer    =   99  'Custom
      Picture         =   "Form4.frx":14318
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False

End Sub


Private Sub Image1_Click()
Me.Visible = False
End Sub

Private Sub Timer1_Timer()
Text4.Text = Text4.Text + 1
If Text4.Text = "60" Then
Form5.Show 1
End If

End Sub
