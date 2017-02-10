VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form DownLoad 
   Caption         =   "正在下载—细胞浏览器"
   ClientHeight    =   1125
   ClientLeft      =   -75
   ClientTop       =   450
   ClientWidth     =   4890
   Icon            =   "FrmDownLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4890
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   120
      Width           =   855
      Begin VB.PictureBox Picture2 
         Height          =   975
         Left            =   -120
         Picture         =   "FrmDownLoad.frx":324A
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   11
         Top             =   -120
         Width           =   975
      End
   End
   Begin VB.TextBox TextOpen 
      Height          =   270
      Left            =   3600
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "托盘"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox downloadName 
      Height          =   270
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   5950
      Begin VB.TextBox TextLocal 
         Height          =   375
         Left            =   1270
         TabIndex        =   2
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox TextURL 
         Height          =   375
         Left            =   1270
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label LabelURL2 
         Caption         =   "保存到："
         Height          =   255
         Left            =   555
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label LabelURL 
         Caption         =   "下载地址："
         Height          =   255
         Left            =   435
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   330
      Left            =   1200
      TabIndex        =   5
      Top             =   645
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   165
      Width           =   2535
   End
End
Attribute VB_Name = "DownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Const NIM_ADD = &H0                     '在任务栏中增加一个图标
Const NIM_DELETE = &H2                  '删除任务栏中的一个图标
Const NIM_MODIFY = &H1                  '修改任务栏中个图标信息
Const NIF_ICON = &H2                    '
Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA结构中uFlags的控制信息
Const NIF_TIP = &H4                     '
Const WM_MOUSEMOVE = &H200              '
Const WM_LBUTTONDBLCLK = &H203          '

Private Type NOTIFYICONDATA
  cbSize As Long                        '该数据结构的大小
  hWnd As Long                          '处理任务栏中图标的窗口句柄
  uID As Long                           '定义的任务栏中图标的标识
  uFlags As Long                        '任务栏图标功能控制，可以是以下值的组合（一般全包括）
                                        'NIF_MESSAGE 表示发送控制消息；
                                        'NIF_ICON表示显示控制栏中的图标；
                                        'NIF_TIP表示任务栏中的图标有动态提示。
  uCallbackMessage As Long '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
  hIcon As Long '任务栏中的图标的控制句柄
  szTip As String * 64 '图标的提示信息
End Type


Dim Tray As NOTIFYICONDATA

Private Sub Command1_Click()
Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hWnd = Me.hWnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = Me.Caption & vbNullChar
Shell_NotifyIcon NIM_ADD, Tray
Me.Hide
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X / 15
If msg = WM_LBUTTONDBLCLK Then
Me.WindowState = 0
Me.Show
Shell_NotifyIcon NIM_DELETE, Tray
End If
End Sub



Private Sub StartDownLoad(ByVal Geturl As String)
    Dim spo%, filename$
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(TextLocal.Text) Then Set f = fso.CreateFolder(TextLocal.Text)
    spo = InStrRev(Geturl, "/")
    filename = Right(Geturl, Len(Geturl) - spo) '获取文件名
    TextLocal.Text = TextLocal.Text & "\" & filename
    Inet1.Execute Geturl, "get"  '开始下载
End Sub







Private Sub Form_Load()
Me.ZOrder
 


TextURL.Text = "http://120343.24la.com.cn/software/cellbrowser.exe"
TextLocal.Text = App.Path
TextOpen.Text = TextLocal.Text

Dim i As Integer
Dim NUm As Integer
NUm = 1
For i = 1 To Len(TextURL.Text)
If Mid(TextURL.Text, i, 1) = "/" Then
NUm = NUm + 1
End If
Next i
l = Split(TextURL.Text, "/", NUm)

downloadName.Text = l(NUm - 1)

StartDownLoad TextURL




End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
msg = MsgBox("正在进行下载任务，确认要退出吗？", vbOKCancel, "提示")
If msg = vbOK Then
On Error Resume Next
 If PB1.Value <> 100 Then Kill (App.Path & "\cellbrowser.exe")
Unload Me
End If
If msg = vbCancel Then
Cancel = 1
End If


End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
Me.Height = 1635
Me.Width = 5120
End If



End Sub

Private Sub Image1_Click()

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
On Error Resume Next
If PB1.Value = 0 Then Command1_Click

    'State = 12 时，用 GetChunk 方法检索服务器的响应。
    Dim vtData() As Byte
    Select Case State
        Case icHostResolvingHost
          Label1.Caption = "正在查询所指定的主机的 IP 地址"
        Case icHostResolved
           Label1.Caption = "成功地找到所指定的主机的 IP 地址"
        Case icConnecting
           Label1.Caption = "正在与主机连接"
        Case icConnected
           Label1.Caption = "已与主机连接成功"
        Case icRequesting
          Label1.Caption = "正在向主机发送请求"
        Case icRequestSent
          Label1.Caption = "发送请求已成功"
        Case icReceivingResponse
          Label1.Caption = "在接收主机的响应"
        Case icResponseReceived
         Label1.Caption = "成功地接收到主机的响应"
        Case icDisconnecting
           Label1.Caption = "正在解除与主机的连接"
        Case icDisconnected
            Label1.Caption = "已成功地与主机解除了连接"
        Case icError
            Label1.Caption = "与主机通讯时出现了错误"
            '出现错误时，返回 ResponseCode 和 ResponseInfo。
            vtData = Inet1.ResponseCode & ":" & Inet1.ResponseInfo
        Case icResponseCompleted ' 12
            Dim bDone As Boolean: bDone = False
            '取得第一个块。
            vtData() = Inet1.GetChunk(1024, 1)
            DoEvents
            

            Open TextLocal.Text For Binary Access Write As #3     '设置保存路径文件后开始保存
                '获取下载文件长度
                If Len(Inet1.GetHeader("Content-Length")) > 0 Then PB1.Max = CLng(Inet1.GetHeader("Content-Length"))
                
                '循环分块下载
                Do While Not bDone
                    Put #3, Loc(3) + 1, vtData()
                    vtData() = Inet1.GetChunk(1024, 1)
                    DoEvents
                    PB1.Value = Loc(3)   '设置进度条长度
                    Label1.Caption = "已下载：" & PB1.Value / PB1.Max * 100 & "%"
                    If Loc(3) >= PB1.Max Then bDone = True
                    Tray.szTip = Me.Caption & vbNullChar
                Loop
            Close #3
     
                        
            Shell App.Path & "\cellbrowser.exe"
            
            End
            
    End Select
End Sub



