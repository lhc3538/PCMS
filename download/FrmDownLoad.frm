VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form DownLoad 
   Caption         =   "�������ء�ϸ�������"
   ClientHeight    =   1125
   ClientLeft      =   -75
   ClientTop       =   450
   ClientWidth     =   4890
   Icon            =   "FrmDownLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4890
   StartUpPosition =   1  '����������
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
      Caption         =   "����"
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
         Caption         =   "���浽��"
         Height          =   255
         Left            =   555
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label LabelURL 
         Caption         =   "���ص�ַ��"
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
         Name            =   "����"
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

Const NIM_ADD = &H0                     '��������������һ��ͼ��
Const NIM_DELETE = &H2                  'ɾ���������е�һ��ͼ��
Const NIM_MODIFY = &H1                  '�޸��������и�ͼ����Ϣ
Const NIF_ICON = &H2                    '
Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA�ṹ��uFlags�Ŀ�����Ϣ
Const NIF_TIP = &H4                     '
Const WM_MOUSEMOVE = &H200              '
Const WM_LBUTTONDBLCLK = &H203          '

Private Type NOTIFYICONDATA
  cbSize As Long                        '�����ݽṹ�Ĵ�С
  hWnd As Long                          '������������ͼ��Ĵ��ھ��
  uID As Long                           '�������������ͼ��ı�ʶ
  uFlags As Long                        '������ͼ�깦�ܿ��ƣ�����������ֵ����ϣ�һ��ȫ������
                                        'NIF_MESSAGE ��ʾ���Ϳ�����Ϣ��
                                        'NIF_ICON��ʾ��ʾ�������е�ͼ�ꣻ
                                        'NIF_TIP��ʾ�������е�ͼ���ж�̬��ʾ��
  uCallbackMessage As Long '������ͼ��ͨ�������û����򽻻���Ϣ���������Ϣ�Ĵ�����hWnd����
  hIcon As Long '�������е�ͼ��Ŀ��ƾ��
  szTip As String * 64 'ͼ�����ʾ��Ϣ
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
    filename = Right(Geturl, Len(Geturl) - spo) '��ȡ�ļ���
    TextLocal.Text = TextLocal.Text & "\" & filename
    Inet1.Execute Geturl, "get"  '��ʼ����
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
msg = MsgBox("���ڽ�����������ȷ��Ҫ�˳���", vbOKCancel, "��ʾ")
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

    'State = 12 ʱ���� GetChunk ������������������Ӧ��
    Dim vtData() As Byte
    Select Case State
        Case icHostResolvingHost
          Label1.Caption = "���ڲ�ѯ��ָ���������� IP ��ַ"
        Case icHostResolved
           Label1.Caption = "�ɹ����ҵ���ָ���������� IP ��ַ"
        Case icConnecting
           Label1.Caption = "��������������"
        Case icConnected
           Label1.Caption = "�����������ӳɹ�"
        Case icRequesting
          Label1.Caption = "������������������"
        Case icRequestSent
          Label1.Caption = "���������ѳɹ�"
        Case icReceivingResponse
          Label1.Caption = "�ڽ�����������Ӧ"
        Case icResponseReceived
         Label1.Caption = "�ɹ��ؽ��յ���������Ӧ"
        Case icDisconnecting
           Label1.Caption = "���ڽ��������������"
        Case icDisconnected
            Label1.Caption = "�ѳɹ������������������"
        Case icError
            Label1.Caption = "������ͨѶʱ�����˴���"
            '���ִ���ʱ������ ResponseCode �� ResponseInfo��
            vtData = Inet1.ResponseCode & ":" & Inet1.ResponseInfo
        Case icResponseCompleted ' 12
            Dim bDone As Boolean: bDone = False
            'ȡ�õ�һ���顣
            vtData() = Inet1.GetChunk(1024, 1)
            DoEvents
            

            Open TextLocal.Text For Binary Access Write As #3     '���ñ���·���ļ���ʼ����
                '��ȡ�����ļ�����
                If Len(Inet1.GetHeader("Content-Length")) > 0 Then PB1.Max = CLng(Inet1.GetHeader("Content-Length"))
                
                'ѭ���ֿ�����
                Do While Not bDone
                    Put #3, Loc(3) + 1, vtData()
                    vtData() = Inet1.GetChunk(1024, 1)
                    DoEvents
                    PB1.Value = Loc(3)   '���ý���������
                    Label1.Caption = "�����أ�" & PB1.Value / PB1.Max * 100 & "%"
                    If Loc(3) >= PB1.Max Then bDone = True
                    Tray.szTip = Me.Caption & vbNullChar
                Loop
            Close #3
     
                        
            Shell App.Path & "\cellbrowser.exe"
            
            End
            
    End Select
End Sub



