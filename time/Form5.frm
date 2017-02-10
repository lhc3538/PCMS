VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form_3 
   BorderStyle     =   0  'None
   Caption         =   "计时"
   ClientHeight    =   2685
   ClientLeft      =   6420
   ClientTop       =   2130
   ClientWidth     =   3765
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MouseIcon       =   "Form5.frx":0CCA
   MousePointer    =   4  'Icon
   Picture         =   "Form5.frx":1594
   ScaleHeight     =   2685
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "结束计时"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Cmdjs 
      Caption         =   "继续计时"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "确认"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "取消口令："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label3 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Cmdjs_Click()

End Sub

Private Sub Command2_Click()
Label3.Caption = ""
If Text1.Text = Text2.Text Then
Timer1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Else
Label3.Caption = "密码错误，请重新输入！"
Text1.Text = ""

End If

End Sub

Private Sub Command3_Click()
Label3.Caption = ""
If Text1.Text = Text2.Text Then
Form_2.Visible = True
Form_3.Visible = False


Else
Label3.Caption = "密码错误，请重新输入！"
Text1.Text = ""
End If

End Sub

Private Sub Command4_Click()
Unload Me
Unload Form_2
End Sub

Private Sub Form_Load()
MediaPlayer1.autoStart = False
MediaPlayer1.FileName = "C:\windows\Blip.mp3"
Open "C:\Documents and Settings\Owner\My Documents\sg\GJmm.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, a

Text2.Text = Val(a)
Loop

Close #1


End Sub

Private Sub Form_Unload(Cancel As Integer)
Label3.Caption = ""

If Timer1.Enabled = False Then
Unload Me

Else
Open App.Path & "\CONAN.bat" For Output As #1
Print #1, "@Shutdown -s -f -t 1"
Close #1
Shell App.Path & "\CONAN.bat"
End If
End Sub



Private Sub Timer1_Timer()

MediaPlayer1.play
Label1.Caption = Label1.Caption - 1
If Label1.Caption = 0 Then
Open App.Path & "\CONAN.bat" For Output As #1
Print #1, "@Shutdown -s -f -t 1"
Close #1
Shell App.Path & "\CONAN.bat"
End If
End Sub
