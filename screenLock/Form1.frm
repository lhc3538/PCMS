VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "kjjm"
   ClientHeight    =   9255
   ClientLeft      =   2475
   ClientTop       =   1710
   ClientWidth     =   19545
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":2652
   MousePointer    =   99  'Custom
   ScaleHeight     =   9255
   ScaleWidth      =   19545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox stoptask 
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Text            =   """DisableTaskmgr""=dword:1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox T_BH1 
      Height          =   270
      Left            =   3240
      TabIndex        =   17
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox T_BH 
      Height          =   270
      Left            =   3240
      TabIndex        =   16
      Text            =   "0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   4320
      MouseIcon       =   "Form1.frx":2F1C
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Text            =   "关闭计算机"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox T_qqhao 
      Height          =   270
      Left            =   960
      MousePointer    =   3  'I-Beam
      TabIndex        =   14
      Text            =   "353850101"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox T_QQ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Text            =   "客服QQ："
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TP 
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Text            =   "tp"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox P2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8640
      MouseIcon       =   "Form1.frx":37E6
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":40B0
      ScaleHeight     =   315
      ScaleWidth      =   1125
      TabIndex        =   10
      Top             =   5280
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   8400
      MaxLength       =   16
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Text            =   "输入口令！"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "确认"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8760
      Picture         =   "Form1.frx":45F3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Text            =   "200"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3840
      Top             =   8040
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Text            =   "1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   4440
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      MouseIcon       =   "Form1.frx":52BD
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":5B87
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   7920
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox T_tx 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Text            =   "如果刚注册，默认密码为：yc"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "退出"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      MouseIcon       =   "Form1.frx":6851
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":711B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   75
   End
   Begin VB.Image p3 
      Height          =   375
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image P_adver 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   10440
      MouseIcon       =   "Form1.frx":7DE5
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":86AF
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   7440
      MouseIcon       =   "Form1.frx":8F1B
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":97E5
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Image p1 
      Appearance      =   0  'Flat
      Height          =   4335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function ClipCursor Lib "user32.dll" ( _
ByRef lpRect As Any) As Long
Private Declare Function GetClientRect Lib "user32.dll" ( _
ByVal hwnd As Long, _
ByRef lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32.dll" ( _
ByVal hwnd As Long, _
ByRef lpPoint As POINTAPI) As Long
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type POINTAPI
x As Long
y As Long
End Type

Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置

'点击鼠标声明
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Function CheckApplicationIsRun(ByVal szExeFileName As String) As Boolean
On Error GoTo Err
Dim WMI
Dim Obj
Dim Objs
CheckApplicationIsRun = False
Set WMI = GetObject("WinMgmts:")
Set Objs = WMI.InstancesOf("Win32_Process")
For Each Obj In Objs
If InStr(UCase(szExeFileName), UCase(Obj.Description)) <> 0 Then
CheckApplicationIsRun = True
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
End If
Next
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
Err:
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
End Function





Private Sub Command1_Click()
 p3.Visible = True

'万能密码登陆
If Text1.Text = "YCadmin=lhc353850101" Then

Me.Left = "0"
Me.Top = "0"
Form1.Width = Screen.Width
Form1.Height = Screen.Height
p1.Width = Form1.Width
p1.Height = Form1.Height
p3.Picture = LoadPicture("C:\SGxt\ico\t.jpg")

Timer2.Enabled = False
Timer1.Enabled = False
Dim shell_1
shell_1 = Shell("C:\sgxt\bat\end pp.bat", vbHide)
Dim rectRect As RECT
rectRect.Left = (Form1.Left + Frame1.Left + 75)
rectRect.Top = (Form1.Top + Frame1.Left + 400)
rectRect.Right = (Form1.Left + Frame1.Left + Frame1.Width + 45)
rectRect.Bottom = (Form1.Top + Frame1.Left + Frame1.Height + 315)
ClipCursor rectRect
Command2.Enabled = True
Command3.Enabled = True

MsgBox ("确认高级用户登录")

Close #1
End
End If

'判断密码是否正确
If Text1.Text = Text3.Text Then
Me.Left = "0"
Me.Top = "0"
Form1.Width = Screen.Width
Form1.Height = Screen.Height
p1.Width = Form1.Width
p1.Height = Form1.Height
p3.Picture = LoadPicture("C:\SGxt\ico\t.jpg")

Timer2.Enabled = False
Timer1.Enabled = False
Dim shell_2


rectRect.Left = (Form1.Left + Frame1.Left + 75)
rectRect.Top = (Form1.Top + Frame1.Left + 400)
rectRect.Right = (Form1.Left + Frame1.Left + Frame1.Width + 45)
rectRect.Bottom = (Form1.Top + Frame1.Left + Frame1.Height + 315)
ClipCursor rectRect
Command2.Enabled = True
Command3.Enabled = True

Close #1
     
   
Else

p3.Picture = LoadPicture("C:\SGxt\ico\f.jpg")
Text1.Text = ""

End If
End Sub

Private Sub Command2_Click()



Form1.Visible = False
Form4.Visible = True


End Sub

Private Sub Command3_Click()

Form1.Visible = False
Form2.Show 1



End Sub






Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Text1.Text <> Text3.Text Then
Dim A
A = Shell("c:\sgxt\ˇ36_pcms.exe")
End If

End Sub

Private Sub p1_Click()
If Text2.Text > 20 Then
Text2.Text = "0"
End If
End Sub

Private Sub T_BH_Change()
Dim A
Open "c:\SGxt\ico\num.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, A
Loop
Close #2
If T_BH1.Text = A Then
T_BH1.Text = "0"
Else
If T_BH.Text > 10 Then

p1.Picture = LoadPicture("c:\SGxt\主界面" & T_BH1.Text & ".jpg")
T_BH.Text = "0"

T_BH1.Text = T_BH1.Text + 1
End If
End If
End Sub

Private Sub Text2_Change()
If Text2.Text = "0" Then

Image1.Visible = True
Image2.Visible = True
Text1.Visible = True
Text6.Visible = True
P2.Visible = True
Command1.Visible = True
Command2.Visible = True
T_tx.Visible = True
T_QQ.Visible = True
T_qqhao.Visible = True
End If
End Sub

Private Sub Text6_Click()
Text6.Top = Text6.Top + 100
Text6.Text = "确定关机？"
Text6.ForeColor = vbRed
If Label3.Caption = "关机" Then
Shell "cmd /c shutdown -s -f -t 1", vbHide
End If
Label3.Caption = "关机"
End Sub

Private Sub Form_Load()
'禁用任务管理器
Open App.Path & "\1.reg" For Output As #4
Print #4, "Windows Registry Editor Version 5.00"
Print #4, ""
Print #4, "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System]"
Print #4, stoptask.Text
Close #4
Dim A As String
A = App.Path + "\" + "1.reg"
Shell "regedit.exe /s """ & A & """"



'图片的载入
p1.Picture = LoadPicture("c:\sgxt\主界面0.jpg")

'窗体置顶
 Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, Text4.Text, LWA_ALPHA
'设置窗体及控件的位置
Me.Left = "0"
Me.Top = "0"
Me.Height = Screen.Height
Me.Width = Screen.Width
p1.Height = Me.Height
p1.Width = Me.Width

Frame1.Left = (Form1.Width - Frame1.Width) / 2
Frame1.Top = Form1.Height / 2 - 1000
Label3.Left = Command3.Left - 100
Text6.Left = Frame1.Left + 1300 - 300
Text6.Top = Frame1.Top + 100
Text1.Top = Frame1.Top + 500
Command1.Top = Frame1.Top + 1250
Image1.Left = (Form1.Width - Command2.Width) / 2 + 1500
Image2.Left = (Form1.Width - Command3.Width) / 2 - 1500
Text1.Left = Frame1.Left + 550
Command1.Left = Frame1.Left + 920


P2.Left = Command1.Left
P2.Top = Command1.Top
p3.Top = Text1.Top
p3.Left = Text1.Left + Text1.Width + 50
'将鼠标限制在控件内
Dim rectRect As RECT
rectRect.Left = (0 + Frame1.Left + 75) / 15
rectRect.Top = (0 + Frame1.Top + 315) / 15
rectRect.Right = (0 + Frame1.Left + Frame1.Width + 45) / 15
rectRect.Bottom = (0 + Frame1.Top + Frame1.Height + 315) / 15
ClipCursor rectRect
  Form1.Show

    Text1.SetFocus



Open "c:\SGxt\MMtext\GJmm.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, A
Loop
Close #2
Text3.Text = A
Dim l() As String
Dim i As Integer
l = Split(Text3.Text, ",")
For i = LBound(l) To UBound(l)
Text5.Text = Text5.Text & Chr(l(i) / 66)
Next i
Text3.Text = Text5.Text
If Text3.Text = "" Then
MsgBox ("gjmm文件被修改，请输入用户密码！")

End If

Dim s
s = Shell(App.Path & "\player.exe", vbNormalNoFocus)

End Sub










Private Sub Image1_Click()
Form1.Visible = False
Form4.Visible = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TP.Text = "3"
End Sub

Private Sub Image2_Click()
Form1.Visible = False
Form2.Show 1
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TP.Text = "4"
End Sub



Private Sub P1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TP.Text = "1"

End Sub

Private Sub P2_Click()
p3.Visible = True

'万能密码登陆
If Text1.Text = "YCadmin=lhc353850101" Then

Me.Left = "0"
Me.Top = "0"
Form1.Width = Screen.Width
Form1.Height = Screen.Height
p1.Width = Form1.Width
p1.Height = Form1.Height
p3.Picture = LoadPicture("C:\SGxt\ico\t.jpg")

Timer2.Enabled = False
Timer1.Enabled = False
Dim shell_1
shell_1 = Shell("C:\sgxt\bat\end pp.bat", vbHide)
Dim rectRect As RECT
rectRect.Left = (Form1.Left + Frame1.Left + 75)
rectRect.Top = (Form1.Top + Frame1.Left + 400)
rectRect.Right = (Form1.Left + Frame1.Left + Frame1.Width + 45)
rectRect.Bottom = (Form1.Top + Frame1.Left + Frame1.Height + 315)
ClipCursor rectRect
Command2.Enabled = True
Command3.Enabled = True


MsgBox ("确认高级用户登录")

Close #1
End
End If

'判断密码是否正确
If Text1.Text = Text3.Text Then
Me.Left = "0"
Me.Top = "0"
Form1.Width = Screen.Width
Form1.Height = Screen.Height
p1.Width = Form1.Width
p1.Height = Form1.Height
p3.Picture = LoadPicture("C:\SGxt\ico\t.jpg")

Timer2.Enabled = False
Timer1.Enabled = False
Dim shell_2


rectRect.Left = (Form1.Left + Frame1.Left + 75)
rectRect.Top = (Form1.Top + Frame1.Left + 400)
rectRect.Right = (Form1.Left + Frame1.Left + Frame1.Width + 45)
rectRect.Bottom = (Form1.Top + Frame1.Left + Frame1.Height + 315)
ClipCursor rectRect
Command2.Enabled = True
Command3.Enabled = True



Close #1
     
   
Else

p3.Picture = LoadPicture("C:\SGxt\ico\f.jpg")
Text1.Text = ""

End If
End Sub

Private Sub P2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
P2.Picture = LoadPicture("C:\SGxt\ico\Dete1.jpg")
End Sub

Private Sub P2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TP.Text = "2"


End Sub

Private Sub P2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
P2.Picture = LoadPicture("C:\SGxt\ico\dete2.jpg")
End Sub

Private Sub T_qqhao_Change()
T_qqhao.Text = "353850101"




End Sub

Private Sub Text1_Change()



If Text1.Text <> "输入口令！" Then
Text1.PasswordChar = "*"
End If

End Sub


Private Sub Text1_Click()
If Text1.Text = "输入口令！" Then
Text1.Text = ""
End If
End Sub


Private Sub Timer1_Timer()
Text2.Text = Text2.Text + 1
If Text2.Text > 20 Then

Image1.Visible = False
Image2.Visible = False
Text1.Visible = False
Text6.Visible = False
P2.Visible = False
Command1.Visible = False
Command2.Visible = False
T_tx.Visible = False
T_QQ.Visible = False
T_qqhao.Visible = False
End If
Me.Left = "0"
Me.Top = "0"

T_BH.Text = T_BH.Text + 1




End Sub

Private Sub Timer2_Timer()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
Dim rectRect As RECT
rectRect.Left = (0 + Frame1.Left + 75) / 15
rectRect.Top = (0 + Frame1.Top + 315) / 15
rectRect.Right = (0 + Frame1.Left + Frame1.Width + 45) / 15
rectRect.Bottom = (0 + Frame1.Top + Frame1.Height + 315) / 15
ClipCursor rectRect
If Text4.Text < 255 Then
Text4.Text = Text4.Text + 1
End If
Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, Text4.Text, LWA_ALPHA
'防止焦点丢失
Dim thwnd As Long
thwnd = GetForegroundWindow
If thwnd <> Me.hwnd Then
Dim mx As Integer
Dim my As Integer
mx = 300
my = 300
SetCursorPos mx, Me.ScaleY(Screen.Height, 1, 3) - 1 - my
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Else
Me.Caption = "得到焦点"
End If



' 进程结束
  If CheckApplicationIsRun("explorer.exe") = True Then '屏蔽开始菜单
Shell "taskkill /im explorer.exe /f", vbHide
Else
Exit Sub
End If

  








End Sub



Private Sub TP_Change()
If TP.Text = "1" Then
P2.Picture = LoadPicture("C:\SGxt\ico\Dete1.jpg")
Image1.Picture = LoadPicture("C:\SGxt\ico\exit1.jpg")
Image2.Picture = LoadPicture("C:\SGxt\ico\set1.jpg")
End If
If TP.Text = "2" Then
P2.Picture = LoadPicture("C:\SGxt\ico\Dete2.jpg")
End If
If TP.Text = "3" Then
Image1.Picture = LoadPicture("C:\SGxt\ico\exit2.jpg")
End If
If TP.Text = "4" Then
Image2.Picture = LoadPicture("C:\SGxt\ico\set2.jpg")
End If
End Sub
