VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "其它"
   ClientHeight    =   1890
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   4230
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4230
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   2400
      MouseIcon       =   "Form4.frx":6912
      MousePointer    =   99  'Custom
      Picture         =   "Form4.frx":6A64
      ScaleHeight     =   915
      ScaleWidth      =   1245
      TabIndex        =   10
      Top             =   360
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton cmdlook 
         BackColor       =   &H00FFFFFF&
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
         MouseIcon       =   "Form4.frx":7AD5
         MousePointer    =   99  'Custom
         Picture         =   "Form4.frx":839F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
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
         MouseIcon       =   "Form4.frx":8B19
         MousePointer    =   99  'Custom
         Picture         =   "Form4.frx":93E3
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "细胞浏览器(新)"
         Height          =   255
         Left            =   1940
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "开机日志"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "计时器"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   5400
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Text            =   """DisableTaskmgr""=dword:00000000"
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "版权所有"
      Height          =   975
      Left            =   3960
      MouseIcon       =   "Form4.frx":9FF8
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   180
      Left            =   3120
      TabIndex        =   2
      Top             =   1200
      Width           =   165
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "问题反馈"
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Form4.frx":A8C2
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Private Sub cmdlook_Click()
l = Shell("c:\sgxt\look.exe", vbHide)
End

End
End Sub

Private Sub Command1_Click()

frmAbout.Show 1
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command4_Click()
l = Shell("c:\sgxt\jsq.exe", vbHide)

End
End Sub

Private Sub Form_Load()
'重新开启任务栏
Shell "explorer.exe "
 If CheckApplicationIsRun("player.exe") = True Then '退出播放器
Shell "taskkill /im player.exe /f", vbHide
Else
Exit Sub
End If
'恢复任务管理器
Open App.Path & "\0.reg" For Output As #4
Print #4, "Windows Registry Editor Version 5.00"
Print #4, ""
Print #4, "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System]"
Print #4, Text2.Text
Close #4
Dim A As String
A = App.Path + "\" + "0.reg"
Shell "regedit.exe /s """ & A & """"

'删除临时文件
Kill (App.Path & "\0.reg")
Kill (App.Path & "\1.reg")

End Sub






Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Form1
Unload Me
End Sub

Private Sub Label3_Click()
juanzhu.Show 1
Me.Hide
End Sub

Private Sub Label4_Click()
Shell "explorer.exe http://120343.24la.com.cn/问题反馈.html", 1
End Sub


Private Sub Picture1_Click()
On Error Resume Next
 Shell App.Path & "\cell.exe", vbNormalFocus

End Sub
