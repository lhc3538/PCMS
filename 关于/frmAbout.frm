VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2625
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
      Height          =   345
      Left            =   4140
      TabIndex        =   1
      Top             =   3075
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   240
      Picture         =   "frmAbout.frx":28EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":45AF
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "˽�˼��������ϵͳ"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "v1.8.5"
      Height          =   225
      Left            =   1560
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "����: һ�а�Ȩ������Ȩ��������������У����е��淢�ֱؾ���                                               �Ŵ�-���"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Top             =   2625
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Visible = False
End Sub

Private Sub cmdSysInfo_Click()
Me.Visible = False
End Sub
