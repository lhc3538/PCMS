VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   3480
   ClientTop       =   1095
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   1440
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2160
      MouseIcon       =   "time提醒.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "time提醒.frx":0152
      Top             =   1320
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "计时器提醒：已到提醒时间！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2


Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Timer1_Timer()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
End Sub
