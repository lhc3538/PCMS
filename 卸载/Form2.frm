VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "私人计算机管理系统 卸载向导"
   ClientHeight    =   5205
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   7695
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7695
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3120
      TabIndex        =   1
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   2160
   End
   Begin VB.Label Label0 
      BackStyle       =   0  'Transparent
      Caption         =   "进行卸载 私人计算机管理系统 "
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image loop1 
      Height          =   300
      Left            =   3600
      Picture         =   "Form2.frx":C84A
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   3240
      MouseIcon       =   "Form2.frx":CB6E
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":CCC0
      Top             =   4750
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   6360
      MouseIcon       =   "Form2.frx":D311
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":D463
      Top             =   4755
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   7680
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Image Image2 
      Height          =   4695
      Left            =   -120
      Picture         =   "Form2.frx":D9A3
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   4800
      MouseIcon       =   "Form2.frx":124D4
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":12626
      Top             =   4750
      Width           =   1125
   End
   Begin VB.Image loop8 
      Height          =   300
      Left            =   6120
      Picture         =   "Form2.frx":12BFA
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image loop7 
      Height          =   300
      Left            =   5760
      Picture         =   "Form2.frx":12F23
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image loop6 
      Height          =   300
      Left            =   5400
      Picture         =   "Form2.frx":1324F
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image loop5 
      Height          =   300
      Left            =   5040
      Picture         =   "Form2.frx":1358A
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image loop4 
      Height          =   300
      Left            =   4680
      Picture         =   "Form2.frx":138C4
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image loop3 
      Height          =   300
      Left            =   4320
      Picture         =   "Form2.frx":13C02
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image loop2 
      Height          =   300
      Left            =   3960
      Picture         =   "Form2.frx":13F3A
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub RecurseTree(CurrPath As String)

    Dim sFileName As String

    Dim newPath As String

    Dim sPath As String

    Static oldPath As String

    sPath = CurrPath & "\"

    sFileName = Dir(sPath, 31) '31的含义∶31=vbNormal+vbReadOnly+vbHidden+vbSystem+vbVolume+vbDirectory

    Do While sFileName <> ""

    If sFileName <> "." And sFileName <> ".." Then

    If GetAttr(sPath & sFileName) And vbDirectory Then '如果是目录和文件夹

    newPath = sPath & sFileName

    RecurseTree newPath

    sFileName = Dir(sPath, 31)

    Else

    SetAttr sPath & sFileName, vbNormal

    Kill (sPath & sFileName)

   
    sFileName = Dir

    End If

    Else

    sFileName = Dir

    End If

    DoEvents

    Loop

    SetAttr CurrPath, vbNormal

    RmDir CurrPath



    End Sub










Private Sub Form_Load()
loop1.Left = 4800
loop1.Top = 2160
loop2.Left = 4800
loop2.Top = 2160
loop3.Left = 4800
loop3.Top = 2160
loop4.Left = 4800
loop4.Top = 2160
loop5.Left = 4800
loop5.Top = 2160
loop6.Left = 4800
loop6.Top = 2160
loop7.Left = 4800
loop7.Top = 2160
loop8.Left = 4800
loop8.Top = 2160
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Me

End Sub

Private Sub Image1_Click()
Timer1.Enabled = True
Image1.Enabled = False
Image3.Enabled = False
Image4.Enabled = False












End Sub

Private Sub Image3_Click()
End
End Sub

Private Sub Image4_Click()
Me.Visible = False
Form1.Visible = True
End Sub

Private Sub Timer1_Timer()

If Text1.Text = 1 Then
Dim strPathName As String
    strPathName = "c:\sgxt\"
  RecurseTree strPathName
loop1.Visible = False
loop2.Visible = True
End If

If Text1.Text = 2 Then
    strPathName = App.Path & "\look\"
  RecurseTree strPathName
loop2.Visible = False
loop3.Visible = True
End If

If Text1.Text = 3 Then
    strPathName = App.Path & "\ico\"
  RecurseTree strPathName
loop3.Visible = False
loop4.Visible = True
End If

If Text1.Text = 4 Then
    strPathName = App.Path & "\time\"
  RecurseTree strPathName
loop4.Visible = False
loop5.Visible = True
End If

If Text1.Text = 5 Then
Kill (App.Path & "\COMCTL32.OCX")
Kill (App.Path & "\COMDLG32.OCX")
Kill (App.Path & "\TABCTL32.OCX")
Kill (App.Path & "\注册.bat")
Kill (App.Path & "\setpage.reg")
loop5.Visible = False
loop6.Visible = True
End If

If Text1.Text = 6 Then
Kill (App.Path & "\num.dat")
Kill (App.Path & "\gs.dat")
Kill (App.Path & "\pp.dat")
Kill (App.Path & "\sj.dat")
loop6.Visible = False
loop7.Visible = True
End If

If Text1.Text = 7 Then
Kill (App.Path & "\kjjm.exe")
Kill (App.Path & "\lock.exe")
Kill (App.Path & "\RET.exe")
Kill (App.Path & "\player.exe")
Kill (App.Path & "\zc.exe")
Kill (App.Path & "\软件升级.exe")
loop7.Visible = False
loop8.Visible = True
End If

If Text1.Text = 8 Then
Kill (App.Path & "\使用说明.txt")
loop8.Visible = False
Me.Visible = False
Form3.Visible = True

End If

Text1.Text = Text1.Text + 1


End Sub
