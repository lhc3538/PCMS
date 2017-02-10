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
   Begin VB.CommandButton CommandExit 
      Caption         =   "退出"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10320
      TabIndex        =   18
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox stoptask 
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Text            =   """DisableTaskmgr""=dword:1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox T_BH1 
      Height          =   270
      Left            =   3240
      TabIndex        =   16
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox T_BH 
      Height          =   270
      Left            =   3240
      TabIndex        =   15
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
      TabIndex        =   14
      Text            =   "关闭计算机"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox T_qqhao 
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   960
      MousePointer    =   3  'I-Beam
      TabIndex        =   13
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
      TabIndex        =   12
      Text            =   "客服QQ："
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TP 
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Text            =   "tp"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   450
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
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
      Enabled         =   0   'False
      Interval        =   100
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
      Height          =   375
      Left            =   7680
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
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
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   11
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
      MouseIcon       =   "Form1.frx":37E6
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":40B0
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
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function ClipCursor Lib "user32.dll" ( _
ByRef lpRect As Any) As Long
Private Declare Function GetClientRect Lib "user32.dll" ( _
ByVal hWnd As Long, _
ByRef lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32.dll" ( _
ByVal hWnd As Long, _
ByRef lpPoint As POINTAPI) As Long
Private Type RECT
left As Long
top As Long
right As Long
bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type

Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置

'点击鼠标声明
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
'获取用户名
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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


p1.Width = Form1.Width
p1.Height = Form1.Height
p3.Picture = LoadPicture("C:\SGxt\ico\t.jpg")

Timer2.Enabled = False
Timer1.Enabled = False
Dim shell_1
shell_1 = Shell("C:\sgxt\bat\end pp.bat", vbHide)
Dim rectRect As RECT
rectRect.left = (Form1.left + Frame1.left + 75)
rectRect.top = (Form1.top + Frame1.left + 400)
rectRect.right = (Form1.left + Frame1.left + Frame1.Width + 45)
rectRect.bottom = (Form1.top + Frame1.left + Frame1.Height + 315)
ClipCursor rectRect
Command2.Enabled = True
Command3.Enabled = True
CommandExit.Enabled = True

MsgBox ("确认高级用户登录")

Close #1
End
End If

'判断密码是否正确
If Text1.Text = Text3.Text Then

p1.Width = Form1.Width
p1.Height = Form1.Height
p3.Picture = LoadPicture("C:\SGxt\ico\t.jpg")

Timer2.Enabled = False
Timer1.Enabled = False
Dim shell_2


rectRect.left = (Form1.left + Frame1.left + 75)
rectRect.top = (Form1.top + Frame1.left + 400)
rectRect.right = (Form1.left + Frame1.left + Frame1.Width + 45)
rectRect.bottom = (Form1.top + Frame1.left + Frame1.Height + 315)
ClipCursor rectRect
Command2.Enabled = True
Command3.Enabled = True
CommandExit.Enabled = True

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






Private Sub CommandExit_Click()
Form1.Visible = False
Form4.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.left = Frame1.left - 300
Command3.top = Frame1.top - 500
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Text1.Text <> Text3.Text Then
Dim A
On Error Resume Next
A = Shell("c:\sgxt\ˇ36_pcms.exe")
If Err Then
Unload Me
MsgBox ("您尚未安装成功，请重新安装！")
Form4.Show
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Text1.Text <> Text3.Text Then
Dim A
On Error Resume Next
A = Shell(App.Path & "\ˇ36_pcms.exe")
If Err Then
Unload Me
MsgBox ("您尚未安装成功，请重新安装！")
Form4.Show
End If
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

Text1.Visible = True
Text6.Visible = True
Command1.Visible = True
Command2.Visible = True
CommandExit.Visible = True
T_tx.Visible = True
T_QQ.Visible = True
T_qqhao.Visible = True
End If
End Sub

Private Sub Text6_Click()
Text6.top = Text6.top + 100
Text6.Text = "确定关机？"
Text6.ForeColor = vbRed
If Label3.Caption = "关机" Then
Shell "cmd /c shutdown -s -f -t 0", vbHide
End If
Label3.Caption = "关机"
End Sub

Private Sub Form_Load()
On Error Resume Next

' 进程结束
  If CheckApplicationIsRun("explorer.exe") = True Then '屏蔽开始菜单
Shell "taskkill /im explorer.exe /f", vbHide
Else
Exit Sub
End If
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
If Err Then GoTo x2

'窗体置顶
 Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, Text4.Text, LWA_ALPHA
'设置窗体及控件的位置
Me.left = "0"
Me.top = "0"
Me.Height = Screen.Height
Me.Width = Screen.Width
p1.Height = Me.Height
p1.Width = Me.Width

Frame1.left = (Form1.Width - Frame1.Width) / 2
Frame1.top = Form1.Height / 2 - 1000
Label3.left = Command3.left - 100
Text6.left = Frame1.left + 1300 - 300
Text6.top = Frame1.top + 100
Text1.top = Frame1.top + 500
Command1.top = Frame1.top + 1250
CommandExit.left = (Form1.Width - Command2.Width) / 2 + 1500
Text1.left = Frame1.left + 550
Command1.left = Frame1.left + 920



p3.top = Text1.top
p3.left = Text1.left + Text1.Width + 50
'将鼠标限制在控件内
Dim rectRect As RECT
rectRect.left = (0 + Frame1.left + 75) / 15
rectRect.top = (0 + Frame1.top + 315) / 15
rectRect.right = (0 + Frame1.left + Frame1.Width + 45) / 15
rectRect.bottom = (0 + Frame1.top + Frame1.Height + 315) / 15
ClipCursor rectRect
  Form1.Show

    Text1.SetFocus

Dim gs As String
Open "c:\SGxt\looktext\gs.dat" For Input As #4
Line Input #4, gs
Close #4
'获取用户名

 Dim st     As String * 100
          Dim pln     As Long
          pln = 99
          GetUserName st, pln
          Dim user As String
        user = left$(st, pln)
Label3.Caption = ":" & Format(Date, "yyyy年mm月dd日") & Format(Now, "hh点mm分ss秒")

gs = gs + 1
Open "c:\SGxt\looktext\gs.dat" For Output As #4
Print #4, gs
Close #4


Dim sj As String

Open "c:\SGxt\looktext\sj.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, sj
Loop
Close #2
sj = sj & "'" & user & Label3.Caption
Open "c:\SGxt\looktext\sj.dat" For Output As #4
Print #4, sj
Close #4


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
s = Shell(App.Path & "\player.exe")

Dim stopcom As Integer
stopcom = 0

 SkinH_Attach
 
If Err Then
x2:
Unload Me
MsgBox ("您尚未安装成功，请重新安装！")
Form4.Show
End If

End Sub





























Private Sub T_qqhao_Change()
T_qqhao.Text = "353850101"




End Sub

Private Sub Text1_Change()



If right(Text1.Text, 5) = "输入口令！" Then
Text1.Text = ""
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
Text1.Visible = False
Text6.Visible = False
Command1.Visible = False
Command2.Visible = False
T_tx.Visible = False
T_QQ.Visible = False
T_qqhao.Visible = False
End If
Me.left = "0"
Me.top = "0"

T_BH.Text = T_BH.Text + 1

Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
Dim rectRect As RECT
rectRect.left = (0 + Frame1.left + 75) / 15
rectRect.top = (0 + Frame1.top + 315) / 15
rectRect.right = (0 + Frame1.left + Frame1.Width + 45) / 15
rectRect.bottom = (0 + Frame1.top + Frame1.Height + 315) / 15
ClipCursor rectRect
If Text4.Text < 255 Then
Text4.Text = Text4.Text + 1
End If
Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, Text4.Text, LWA_ALPHA
'防止焦点丢失
Dim thwnd As Long
thwnd = GetForegroundWindow
If thwnd <> Me.hWnd Then
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



